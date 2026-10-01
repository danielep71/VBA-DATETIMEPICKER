#!/usr/bin/env python3
"""Validate a DateTimePicker evidence record and exact local package hashes.

This checks consistency of maintainer assertions, not their authenticity. It
does not run Excel, build packages, publish releases, or close milestones.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
from pathlib import Path

from _gatelib import git_text, write_json

MANUAL = ("provider-ownership", "clock-lifecycle", "grid-ownership",
          "application-state", "packaged-entry-paths", "ribbon-image-provenance")
STATIC_CHECKS = {"tool-tests", "check_vba_conditionals-fixtures", "check_vba_jumps-fixtures",
                 "check_committed_whitespace-fixtures", "check_source", "check_vba_conditionals",
                 "check_vba_jumps", "check_committed_whitespace", "labels-sync", "labels-drift"}


def require(condition: object, message: str) -> None:
    if not condition:
        raise ValueError(message)


def read_object(path: Path) -> dict:
    value = json.loads(path.read_text(encoding="utf-8"))
    require(isinstance(value, dict), f"{path}: expected an object")
    return value


def validate(root: Path, record: dict, assets_dir: Path, static: dict,
             require_tag: bool = False) -> dict:
    """Fail closed on missing evidence, stale identity, failures or changed bytes."""
    require(record.get("schema_version") == 1, "Unsupported evidence schema")
    sha = record.get("candidate_sha")
    require(isinstance(sha, str) and re.fullmatch(r"[0-9a-f]{40}", sha), "Full candidate SHA required")
    resolved = git_text(root, "rev-parse", "--verify", f"{sha}^{{commit}}", check=True).stdout.strip()
    require(resolved == sha, "Candidate must identify a commit")
    main = git_text(root, "rev-parse", "--verify", "refs/remotes/origin/main^{commit}", check=True).stdout.strip()
    ancestry = git_text(root, "merge-base", "--is-ancestor", sha, main)
    require(ancestry.returncode == 0,
            "Candidate must be reachable from fetched origin/main; run git fetch origin main")
    version = git_text(root, "show", f"{sha}:VERSION", check=True).stdout.strip()
    require(re.fullmatch(r"(?:0|[1-9]\d*)\.(?:0|[1-9]\d*)\.(?:0|[1-9]\d*)", version), "Invalid candidate VERSION")
    require(record.get("version") == version and record.get("tag") == f"v{version}", "Version/tag mismatch")
    changelog = git_text(root, "show", f"{sha}:CHANGELOG.md", check=True).stdout
    require(re.search(rf"^## \[{re.escape(version)}\] - \d{{4}}-\d{{2}}-\d{{2}}$", changelog, re.M),
            "Candidate needs a dated changelog section")
    if require_tag:
        ref = "refs/tags/" + record["tag"]
        require(git_text(root, "cat-file", "-t", ref, check=True).stdout.strip() == "tag", "Annotated tag required")
        require(git_text(root, "rev-parse", ref + "^{commit}", check=True).stdout.strip() == sha, "Tag targets another commit")
    require(static.get("status") == "pass" and static.get("candidate_sha") == sha
            and static.get("dirty") is False and static.get("mode") == "committed",
            "Static report must PASS on the clean exact committed candidate")
    checks = static.get("checks")
    require(isinstance(checks, list) and bool(checks)
            and all(isinstance(c, dict) and c.get("exit_code") == 0 for c in checks),
            "Static report has no successful checks")
    require(len(checks) == len(STATIC_CHECKS) and {c.get("name") for c in checks} == STATIC_CHECKS,
            "Static report is incomplete or contains duplicate checks")
    assets = record.get("assets")
    require(isinstance(assets, list) and len(assets) == 2, "Exactly two package records required")
    expected = {"xlam": f"DATETIMEPICKER.v{version}.xlam",
                "xlsm": f"DATETIMEPICKER-demo-v{version}.xlsm"}
    require(all(isinstance(a, dict) for a in assets), "Invalid package records")
    require({a.get("host") for a in assets} == set(expected), "Both xlsm and xlam hosts required")
    for asset in assets:
        host = asset["host"]
        require(asset.get("filename") == expected[host], f"Unexpected {host} package name")
        require(asset.get("candidate_sha") == sha, f"{host}: stale package source identity")
        path = assets_dir / expected[host]
        require(path.is_file() and not path.is_symlink(), f"{host}: package is missing or a symlink")
        require(type(asset.get("size")) is int and asset["size"] > 0
                and path.stat().st_size == asset["size"], f"{host}: size mismatch")
        with path.open("rb") as handle:
            digest = hashlib.sha256()
            for block in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(block)
        require(asset.get("sha256") == digest.hexdigest(), f"{host}: SHA-256 mismatch")
        require(asset.get("compile") == "PASS", f"{host}: compile has not passed")
        environment = asset.get("environment")
        require(isinstance(environment, dict), f"{host}: environment required")
        for field in ("excel", "windows", "bitness", "scaling", "monitors", "load_mode", "other_addins"):
            require(isinstance(environment.get(field), str) and environment[field].strip(),
                    f"{host}: environment.{field} required")
        for pack in ("standard", "ui_smoke"):
            result = asset.get(pack)
            require(isinstance(result, dict), f"{host}: {pack} result required")
            require(result.get("state") == "PASS", f"{host}: {pack} did not PASS")
            for count in ("run", "passed", "failed", "cleanup_failures", "expected_run", "suites", "expected_suites"):
                require(type(result.get(count)) is int, f"{host}: {pack}.{count} must be an integer")
            require(result["run"] == result["passed"] == result["expected_run"] > 0
                    and result["failed"] == result["cleanup_failures"] == 0
                    and result["suites"] == result["expected_suites"] > 0,
                    f"{host}: {pack} is failed, incomplete or has cleanup failures")
    manual = record.get("manual")
    require(isinstance(manual, dict) and set(manual) == set(MANUAL), "All manual matrix records required")
    limitations = []
    for name in MANUAL:
        item = manual[name]
        require(isinstance(item, dict), f"{name}: invalid manual record")
        require(item.get("status") in {"PASS", "NOT RUN"}, f"{name}: unresolved failure or unknown status")
        require(isinstance(item.get("detail"), str) and item["detail"].strip(), f"{name}: describe evidence or gap")
        if item["status"] == "NOT RUN":
            require(isinstance(item.get("accepted_by"), str) and item["accepted_by"].strip(),
                    f"{name}: gap has not been explicitly accepted")
            limitations.append({"check": name, "detail": item["detail"], "accepted_by": item["accepted_by"]})
    return {"schema_version": 1, "candidate_sha": sha, "tag": record["tag"],
            "verified_main_sha": main,
            "status": "pass_with_limitations" if limitations else "pass",
            "accepted_limitations": limitations, "tag_verified": require_tag,
            "scope": "Record consistency and local hashes; Excel assertions are maintainer-supplied."}


def prepare(root: Path, candidate: str, assets_dir: Path) -> dict:
    """Prepare an explicitly untested record; fill results after testing these bytes."""
    require(re.fullmatch(r"[0-9a-f]{40}", candidate), "Full candidate SHA required")
    version = git_text(root, "show", f"{candidate}:VERSION", check=True).stdout.strip()
    require(re.fullmatch(r"\d+\.\d+\.\d+", version), "Invalid candidate VERSION")
    assets = []
    for host, filename in (("xlam", f"DATETIMEPICKER.v{version}.xlam"),
                           ("xlsm", f"DATETIMEPICKER-demo-v{version}.xlsm")):
        path = assets_dir / filename
        digest = hashlib.sha256()
        with path.open("rb") as handle:
            for block in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(block)
        assets.append({"host": host, "filename": filename, "size": path.stat().st_size,
                       "sha256": digest.hexdigest(), "candidate_sha": candidate,
                       "compile": "NOT RUN", "environment": {field: "" for field in
                       ("excel", "windows", "bitness", "scaling", "monitors", "load_mode", "other_addins")},
                       **{pack: {"state": "NOT RUN", "run": 0, "passed": 0, "failed": 0,
                                 "cleanup_failures": 0, "expected_run": 0, "suites": 0,
                                 "expected_suites": 0} for pack in ("standard", "ui_smoke")}})
    return {"schema_version": 1, "candidate_sha": candidate, "version": version,
            "tag": f"v{version}", "assets": assets,
            "manual": {name: {"status": "NOT RUN", "detail": "", "accepted_by": ""} for name in MANUAL}}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path.cwd())
    parser.add_argument("--evidence", type=Path)
    parser.add_argument("--assets-dir", type=Path, required=True)
    parser.add_argument("--static-report", type=Path)
    parser.add_argument("--prepare", action="store_true")
    parser.add_argument("--candidate-sha")
    parser.add_argument("--require-tag", action="store_true")
    parser.add_argument("--output", type=Path, default=Path("test-results/release-check.json"))
    args = parser.parse_args()
    if args.prepare and not args.candidate_sha:
        parser.error("--prepare requires --candidate-sha")
    if not args.prepare and (not args.evidence or not args.static_report):
        parser.error("validation requires --evidence and --static-report")
    try:
        if args.prepare:
            require(not args.output.exists(), "Refusing to overwrite an existing evidence record")
            write_json(args.output, prepare(args.root, args.candidate_sha, args.assets_dir))
            print(f"Prepared {args.output}; Excel results are NOT RUN")
            return 0
        report = validate(args.root, read_object(args.evidence), args.assets_dir,
                          read_object(args.static_report), args.require_tag)
    except (ValueError, OSError) as error:
        report = {"status": "fail", "error": str(error)}
    except subprocess.CalledProcessError as error:
        report = {"status": "fail", "error": f"Git identity lookup failed: {error.stderr.strip()}"}
    # A failed preparation must not overwrite an existing certification record.
    if not args.prepare:
        require(args.output.resolve() != args.evidence.resolve(), "Output must not overwrite evidence")
        require(args.output.resolve() != args.static_report.resolve(), "Output must not overwrite static report")
        write_json(args.output, report)
    print(json.dumps(report, indent=2))
    return 1 if report["status"] == "fail" else 0


if __name__ == "__main__":
    raise SystemExit(main())
