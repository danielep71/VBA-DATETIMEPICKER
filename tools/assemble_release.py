#!/usr/bin/env python3
"""Assemble release-evidence.json, the tag message and the release notes (#222).

Every value comes from a recorded file, never from typed numbers:

- the evidence folders of tools/Certify-Excel.ps1 (two source runs and one
  package run per package) and of the package builds;
- the candidate's own VERSION, CHANGELOG.md and regression harness, read from
  Git at the candidate SHA;
- the exact-candidate static-checks.json;
- the operator's answers to the six manual checks (manual-checks.json);
- the packages in the release-assets folder, hashed here.

The tool cross-checks them: one candidate SHA everywhere, clean working trees,
every run PASS, each package's hash equal across its build, its package run and
the asset, and each pack's totals equal to the source run of the same commit.
It then validates the record with check_release.validate and writes four files
into a new output folder. It changes nothing else and never publishes.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
from pathlib import Path
from typing import Any, cast

import check_release
from _gatelib import git_text, write_text

REPOSITORY_URL = "https://github.com/danielep71/VBA-DATETIMEPICKER"
CERTIFY = "tools/Certify-Excel.ps1"
BUILDERS = {"xlam": "tools/Build-Addin.ps1", "xlsm": "tools/Build-Demo.ps1"}
PACKS = {"standard": "TST_DP_RunAll", "ui_smoke": "TST_DP_RunAll_WithUISmoke"}
UI_SMOKE_SUITE = "UISmoke"
HARNESS = "test/M_cDP_Test.bas"
REQUIRED_KNOWN = ("excel", "windows", "bitness")
NOTES_LIMIT = 120_000
OUTPUTS = ("release-evidence.json", "release-check.json", "tag-message.txt", "release-notes.md")


def require(condition: object, message: str) -> None:
    if not condition:
        raise ValueError(message)


def sha256_of(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def as_object(value: object, where: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{where}: expected an object")
    return cast(dict[str, Any], value)


def load_summary(folder: Path, label: str) -> dict[str, Any]:
    path = folder / "summary.json"
    require(path.is_file(), f"{label}: {path} not found")
    try:
        value: object = json.loads(path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as error:
        raise ValueError(f"{label}: summary.json is not valid JSON ({error})") from error
    return as_object(value, label)


def check_common(summary: dict[str, Any], label: str, tool: str, sha: str) -> None:
    """The fields every tools/*.ps1 evidence folder shares."""
    require(summary.get("tool") == tool, f"{label}: written by {summary.get('tool')!r}, not {tool}")
    require(summary.get("candidate_sha") == sha,
            f"{label}: candidate {summary.get('candidate_sha')!r} is not {sha}")
    require(summary.get("working_tree_clean") is True, f"{label}: the working tree was not clean")
    reasons = summary.get("reasons") or []
    require(summary.get("result") == "PASS" and not reasons,
            f"{label}: result {summary.get('result')!r}" + (f" ({'; '.join(map(str, reasons))})" if reasons else ""))


def count(record: dict[str, Any], field: str, label: str) -> int:
    value = record.get(field)
    require(type(value) is int, f"{label}: {field} is not a whole number")
    return cast(int, value)


def checked_harness(harness: object, label: str) -> dict[str, int]:
    """The counts of one harness record that passed completely."""
    record = as_object(harness, label)
    require(record.get("status") == "RAN", f"{label}: harness {record.get('status')!r}")
    require(record.get("state") == "PASS", f"{label}: state {record.get('state')!r}")
    counts = {field: count(record, field, label) for field in
              ("run", "passed", "failed", "cleanup_failures", "suites_dispatched", "suites_completed")}
    require(counts["run"] == counts["passed"] > 0 and counts["failed"] == counts["cleanup_failures"] == 0,
            f"{label}: Run={counts['run']}, Passed={counts['passed']}, Failed={counts['failed']}, "
            f"CleanupFailures={counts['cleanup_failures']}")
    require(counts["suites_completed"] == counts["suites_dispatched"] > 0,
            f"{label}: {counts['suites_completed']} of {counts['suites_dispatched']} suites completed")
    return counts


def expected_suites(root: Path, sha: str) -> dict[str, int]:
    """Suites each pack dispatches, counted in the candidate's harness source."""
    source = git_text(root, "show", f"{sha}:{HARNESS}", check=True).stdout
    names = re.findall(r'^[ \t]*TST_DP_RunSuiteSafe "(\w+)"', source, flags=re.M)
    require(names and names.count(UI_SMOKE_SUITE) == 1,
            f"{HARNESS} at {sha[:12]}: expected suite dispatches with one {UI_SMOKE_SUITE}")
    return {"standard": len(names) - 1, "ui_smoke": len(names)}


def source_totals(folders: dict[str, Path], sha: str, suites: dict[str, int]) -> dict[str, int]:
    """Assertion totals of the source runs, which the package runs must equal."""
    totals = {}
    for pack, entry in PACKS.items():
        label = f"source {pack} run"
        summary = load_summary(folders[pack], label)
        check_common(summary, label, CERTIFY, sha)
        require(summary.get("mode") == "source", f"{label}: mode {summary.get('mode')!r}, not source")
        require(summary.get("entry_point") == entry, f"{label}: entry point {summary.get('entry_point')!r}")
        counts = checked_harness(summary.get("harness"), label)
        require(counts["suites_dispatched"] == suites[pack],
                f"{label}: {counts['suites_dispatched']} suites dispatched, the harness has {suites[pack]}")
        totals[pack] = counts["run"]
    return totals


def package_asset(host: str, build_dir: Path, run_dir: Path, assets_dir: Path, version: str, sha: str,
                  suites: dict[str, int], totals: dict[str, int]) -> dict[str, Any]:
    """One asset record of release-evidence.json, checked against its build and package run."""
    filename = (f"DATETIMEPICKER.v{version}.xlam" if host == "xlam"
                else f"DATETIMEPICKER-demo-v{version}.xlsm")
    path = assets_dir / filename
    require(path.is_file() and not path.is_symlink(), f"{host}: {path} is missing or a symlink")
    digest, size = sha256_of(path), path.stat().st_size

    label = f"{host} build"
    build = load_summary(build_dir, label)
    check_common(build, label, BUILDERS[host], sha)
    built = as_object(build.get("package"), label)
    require(built.get("sha256") == digest and built.get("bytes") == size,
            f"{label}: the built package is not {filename} in {assets_dir}")
    require(built.get("sha256_after_reopen") == digest, f"{label}: the package changed after reopening")
    require(as_object(build.get("metadata"), label).get("status") == "PASS",
            f"{label}: build-machine metadata was not removed")

    label = f"{host} package run"
    run = load_summary(run_dir, label)
    check_common(run, label, CERTIFY, sha)
    require(run.get("mode") == "package", f"{label}: mode {run.get('mode')!r}, not package")
    tested = as_object(run.get("package"), label)
    require(tested.get("host") == host, f"{label}: tested a {tested.get('host')!r} package")
    require(tested.get("sha256") == digest and tested.get("bytes") == size,
            f"{label}: the tested package is not {filename}")
    packs = as_object(run.get("packs"), label)
    record: dict[str, Any] = {"host": host, "filename": filename, "size": size, "sha256": digest,
                              "candidate_sha": sha}
    for pack, entry in PACKS.items():
        pack_label = f"{label}, {pack}"
        result = as_object(packs.get(pack), pack_label)
        require(result.get("entry_point") == entry, f"{pack_label}: entry point {result.get('entry_point')!r}")
        require(as_object(result.get("compile"), pack_label).get("status") == "PASS",
                f"{pack_label}: compile did not pass")
        require(result.get("sha256_after") == digest, f"{pack_label}: the package changed during the pack")
        counts = checked_harness(result.get("harness"), pack_label)
        record[pack] = {"state": "PASS", "run": counts["run"], "passed": counts["passed"],
                        "failed": counts["failed"], "cleanup_failures": counts["cleanup_failures"],
                        "expected_run": totals[pack], "suites": counts["suites_completed"],
                        "expected_suites": suites[pack]}
    record["compile"] = "PASS"
    found = as_object(run.get("host"), label)
    environment = {"excel": found.get("excel_build"), "windows": found.get("os"),
                   "bitness": found.get("office_bitness"), "scaling": found.get("scaling"),
                   "monitors": found.get("monitors"), "load_mode": found.get("load_mode"),
                   "other_addins": found.get("other_addins")}
    for field, value in environment.items():
        require(isinstance(value, str) and value.strip(), f"{label}: host has no {field}")
        if field in REQUIRED_KNOWN:
            require(value != "unknown", f"{label}: {field} is unknown")
    record["environment"] = environment
    return record


def read_manual(path: Path) -> dict[str, Any]:
    """The six manual records as the operator answered them."""
    value: object = json.loads(path.read_text(encoding="utf-8-sig"))
    manual = as_object(value, str(path))
    require(set(manual) == set(check_release.MANUAL),
            f"{path}: expected exactly {', '.join(check_release.MANUAL)}")
    return manual


def changelog_section(root: Path, sha: str, version: str) -> tuple[str, str]:
    """The dated CHANGELOG section of the version, and its date."""
    text = git_text(root, "show", f"{sha}:CHANGELOG.md", check=True).stdout
    heading = re.search(rf"^## \[{re.escape(version)}\] - (\d{{4}}-\d{{2}}-\d{{2}})[ \t]*$", text, flags=re.M)
    require(heading, f"CHANGELOG.md at {sha[:12]} has no dated [{version}] section")
    heading = cast(re.Match[str], heading)
    rest = text[heading.end():]
    following = re.search(r"^## \[", rest, flags=re.M)
    body = rest[:following.start()] if following else rest
    body = re.sub(r"^<a id=\"[^\"]*\"></a>\s*$", "", body, flags=re.M)
    return body.strip().rstrip("-").strip(), heading.group(1)


def absolute_links(text: str, tag: str) -> str:
    """Repository-relative links resolved at the tag, as a release page needs."""
    def resolve(match: re.Match[str]) -> str:
        target = match.group(2)
        if re.match(r"[a-z][a-z0-9+.-]*:|#|/", target, flags=re.I):
            return match.group(0)
        return f"{match.group(1)}({REPOSITORY_URL}/blob/{tag}/{target})"
    return re.sub(r"(\]\s?)\(([^)\s]+)\)", resolve, text)


def previous_tag(root: Path, sha: str, tag: str) -> str | None:
    listed = git_text(root, "tag", "--merged", sha, "--sort=-v:refname", "--list", "v*", check=True).stdout
    for name in listed.split():
        if name != tag and re.fullmatch(r"v\d+\.\d+\.\d+", name):
            return name
    return None


def deviations(manual: dict[str, Any]) -> list[str]:
    return [f"{name} not run ({item['detail']}; accepted by {item['accepted_by']})"
            for name, item in manual.items() if item.get("status") == "NOT RUN"]


def tag_message(record: dict[str, Any], title: str, refs: list[int], evidence_issue: int) -> str:
    """The annotated tag's message, in the RELEASING section 9 template."""
    hosts = {asset["host"]: asset for asset in record["assets"]}
    lines = [f"{record['tag']} - {title}", "", f"Certified at {record['candidate_sha']}."]
    for host, label in (("xlsm", "Embedded .xlsm"), ("xlam", "Packaged .xlam")):
        asset = hosts[host]
        lines.append(f"{label}: standard {asset['standard']['passed']}/{asset['standard']['run']}, "
                     f"UI smoke {asset['ui_smoke']['passed']}/{asset['ui_smoke']['run']}, "
                     "Failed=0, CleanupFailures=0.")
    if refs:
        lines.append("Refs " + ", ".join(f"#{n}" for n in refs) + ".")
    found = deviations(record["manual"])
    lines.append("Deviations: " + ("; ".join(found) if found else "none") + ".")
    lines.append(f"Evidence: #{evidence_issue}.")
    return "\n".join(lines) + "\n"


def environment_lines(hosts: dict[str, dict[str, Any]]) -> list[str]:
    """The test environment, once when both packages ran in the same one."""
    def describe(environment: dict[str, str]) -> str:
        return (f"Excel {environment['excel']}, {environment['bitness']}; {environment['windows']}; "
                f"{environment['monitors']} monitor(s) at {environment['scaling']}")
    shown = {host: describe(hosts[host]["environment"]) for host in ("xlam", "xlsm")}
    if shown["xlam"] == shown["xlsm"]:
        return [f"Environment: {shown['xlam']}."]
    return [f"Environment of the `.xlam` run: {shown['xlam']}.", "",
            f"Environment of the `.xlsm` run: {shown['xlsm']}."]


def release_notes(root: Path, record: dict[str, Any], title: str, evidence_issue: int,
                  manual_pdf: dict[str, Any]) -> str:
    """The GitHub Release body: changes, assets with SHA-256, and certification."""
    sha, tag, version = record["candidate_sha"], record["tag"], record["version"]
    changes, date = changelog_section(root, sha, version)
    hosts = {asset["host"]: asset for asset in record["assets"]}
    lines = [f"# {tag} — {title}", "",
             f"Released {date}. Certified at [`{sha}`]({REPOSITORY_URL}/commit/{sha}); "
             f"the annotated tag `{tag}` points to that commit.", "",
             "## Changes", "", absolute_links(changes, tag), "",
             "## Downloads", "",
             "| File | Bytes | SHA-256 |", "| --- | ---: | --- |"]
    for item in (hosts["xlam"], hosts["xlsm"], manual_pdf):
        lines.append(f"| `{item['filename']}` | {item['size']} | `{item['sha256']}` |")
    lines += ["", "`release-evidence.json` is the machine-readable certification record. "
              "A workflow checks it against the published files after publication.", "",
              "The `.xlam` and `.xlsm` files contain VBA macros. Download them only from this page and "
              "compare the SHA-256 before you enable macros:", "",
              "```powershell", f"Get-FileHash -Algorithm SHA256 .\\{hosts['xlam']['filename']}", "```", "",
              "## Certification", "",
              "| Package | Pack | Assertions | Suites | Failed | Cleanup failures |",
              "| --- | --- | ---: | ---: | ---: | ---: |"]
    for host, label in (("xlam", "Add-in `.xlam`"), ("xlsm", "Demo `.xlsm`")):
        for pack, name in (("standard", "Standard"), ("ui_smoke", "With UI smoke")):
            result = hosts[host][pack]
            lines.append(f"| {label} | {name} | {result['passed']}/{result['run']} | "
                         f"{result['suites']}/{result['expected_suites']} | {result['failed']} | "
                         f"{result['cleanup_failures']} |")
    lines += ["", "Each package was compiled and tested after its build, opened read-only, and its "
              "SHA-256 was unchanged after every run. Each pack's totals equal the source run of the same "
              "commit.", "",
              *environment_lines(hosts), "",
              "| Manual check | Result | Detail |", "| --- | --- | --- |"]
    for name in check_release.MANUAL:
        item = record["manual"][name]
        detail = str(item["detail"]).replace("|", "\\|").replace("\n", " ")
        lines.append(f"| `{name}` | {item['status']} | {detail} |")
    found = deviations(record["manual"])
    lines += ["", "**Deviations:** " + ("; ".join(found) if found else "none") + ".",
              f"**Certification record:** #{evidence_issue}."]
    previous = previous_tag(root, sha, tag)
    if previous:
        lines += ["", f"**Full changelog:** {REPOSITORY_URL}/compare/{previous}...{tag}"]
    text = "\n".join(lines) + "\n"
    require(len(text) <= NOTES_LIMIT, f"The release notes are {len(text)} characters; GitHub allows 125000")
    return text


def assemble(args: argparse.Namespace) -> dict[str, str]:
    """Build and validate every output; nothing is written when a check fails."""
    root: Path = args.root.resolve()
    sha: str = args.candidate_sha
    require(re.fullmatch(r"[0-9a-f]{40}", sha), "Full candidate SHA required")
    version = git_text(root, "show", f"{sha}:VERSION", check=True).stdout.strip()
    require(re.fullmatch(r"\d+\.\d+\.\d+", version), f"Invalid VERSION at {sha[:12]}")
    require("\n" not in args.title and args.title.strip(), "The title must be one non-empty line")
    suites = expected_suites(root, sha)
    totals = source_totals({"standard": args.source_standard, "ui_smoke": args.source_ui_smoke}, sha, suites)
    assets = []
    for host, build_dir, run_dir in (("xlam", args.addin_build, args.addin_run),
                                     ("xlsm", args.demo_build, args.demo_run)):
        assets.append(package_asset(host, build_dir, run_dir, args.assets_dir, version, sha, suites, totals))
    pdf = args.assets_dir / f"DATETIMEPICKER-user-manual-v{version}.pdf"
    require(pdf.is_file() and not pdf.is_symlink(), f"{pdf} is missing or a symlink")
    manual_pdf = {"filename": pdf.name, "size": pdf.stat().st_size, "sha256": sha256_of(pdf)}
    record: dict[str, Any] = {"schema_version": 1, "candidate_sha": sha, "version": version,
                              "tag": f"v{version}", "assets": assets, "documents": [manual_pdf],
                              "manual": read_manual(args.manual)}
    static = check_release.read_object(args.static_report)
    report = check_release.validate(root, record, args.assets_dir, static)
    message = tag_message(record, args.title.strip(), sorted(set(args.refs)), args.evidence_issue)
    notes = release_notes(root, record, args.title.strip(), args.evidence_issue, manual_pdf)
    return {"release-evidence.json": json.dumps(record, indent=2, sort_keys=True) + "\n",
            "release-check.json": json.dumps(report, indent=2, sort_keys=True) + "\n",
            "tag-message.txt": message, "release-notes.md": notes}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--candidate-sha", required=True)
    parser.add_argument("--assets-dir", type=Path, required=True)
    parser.add_argument("--static-report", type=Path, required=True)
    parser.add_argument("--source-standard", type=Path, required=True, help="Certify-Excel source run, standard")
    parser.add_argument("--source-ui-smoke", type=Path, required=True, help="Certify-Excel source run, UI smoke")
    parser.add_argument("--addin-build", type=Path, required=True)
    parser.add_argument("--addin-run", type=Path, required=True)
    parser.add_argument("--demo-build", type=Path, required=True)
    parser.add_argument("--demo-run", type=Path, required=True)
    parser.add_argument("--manual", type=Path, required=True, help="manual-checks.json with the six records")
    parser.add_argument("--title", required=True, help="Short release title for the tag and the notes")
    parser.add_argument("--refs", type=int, nargs="*", default=[], help="Issue numbers for the tag message")
    parser.add_argument("--evidence-issue", type=int, required=True, help="Certification issue number")
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    output: Path = args.output_dir
    try:
        require(not output.exists() or not any(output.iterdir()),
                f"{output} is not empty; evidence is never overwritten")
        files = assemble(args)
    except (ValueError, OSError, json.JSONDecodeError) as error:
        print(f"FAIL: {error}")
        return 1
    except subprocess.CalledProcessError as error:
        print(f"FAIL: Git could not read the candidate: {error.stderr.strip()}")
        return 1
    for name in OUTPUTS:
        write_text(output / name, files[name])
    report = json.loads(files["release-check.json"])
    print(f"{report['status'].upper()}: {output}")
    for item in report["accepted_limitations"]:
        print(f"  accepted limitation: {item['check']} ({item['accepted_by']})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
