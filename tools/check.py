#!/usr/bin/env python3
"""Run the same portable source checks locally and in CI; no Office required."""
from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

from _gatelib import git_text, write_json, write_text


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--output-dir", type=Path, default=Path("test-results"))
    parser.add_argument("--ci", action="store_true", help="Check committed whitespace, not local edits")
    parser.add_argument("--base", help="PR base revision for committed whitespace")
    args = parser.parse_args()
    root = args.root.resolve()
    output = args.output_dir.resolve()
    output.mkdir(parents=True, exist_ok=True)
    sha = git_text(root, "rev-parse", "HEAD", check=True).stdout.strip()
    dirty = bool(git_text(root, "status", "--porcelain", "--untracked-files=normal", check=True).stdout)
    commands = [
        ("tool-tests", [sys.executable, "-m", "unittest", "discover", "-s", "tools", "-p", "test_*.py"]),
    ]
    for name in ("check_vba_conditionals", "check_vba_jumps", "check_committed_whitespace"):
        commands.append((name + "-fixtures", [sys.executable, f"tools/{name}.py", "--self-test"]))
    for name in ("check_source", "check_vba_conditionals", "check_vba_jumps", "check_committed_whitespace"):
        command = [sys.executable, f"tools/{name}.py", "--root", str(root),
                   "--output", str(output / f"{name}.json")]
        if name == "check_committed_whitespace":
            command += ["--mode", "committed" if args.ci else "working-tree"]
            if args.ci and args.base:
                command += ["--base", args.base]
        commands.append((name, command))
    for name in ("labels-sync", "labels-drift"):
        commands.append((name, ["node", f".github/scripts/{name}.mjs", "--policy",
                                ".github/labels-policy.json", "--self-test"]))
    results = []
    for name, command in commands:
        try:
            result = subprocess.run(command, cwd=root, capture_output=True, text=True, check=False)
            code, log = result.returncode, result.stdout + result.stderr
        except OSError as error:
            code, log = 2, str(error)
        write_text(output / f"{name}.log", log)
        results.append({"name": name, "exit_code": code})
        print(f"{name}: {'PASS' if code == 0 else 'FAIL'}", flush=True)
        if code:
            print(log, flush=True)
    passed = all(r["exit_code"] == 0 for r in results) and (not args.ci or not dirty)
    report = {"schema_version": 1, "candidate_sha": sha, "dirty": dirty,
              "mode": "committed" if args.ci else "working-tree",
              "status": "pass" if passed else "fail", "checks": results,
              "scope": "Static source and tooling checks only; Excel has not been executed."}
    write_json(output / "static-checks.json", report)
    summary = "# Static checks\n\n" + f"Status: **{report['status'].upper()}**\n\n"
    summary += f"Candidate: `{sha}`; dirty: `{dirty}`\n\n" + report["scope"] + "\n"
    summary += "\n".join(f"- {r['name']}: {'PASS' if r['exit_code'] == 0 else 'FAIL'}" for r in results) + "\n"
    write_text(output / "static-checks.md", summary)
    return 0 if passed else 1


if __name__ == "__main__":
    raise SystemExit(main())
