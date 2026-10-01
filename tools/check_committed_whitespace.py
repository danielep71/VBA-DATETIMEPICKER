#!/usr/bin/env python3
"""Validate committed or local Git whitespace without mutating the repository."""

from __future__ import annotations

import argparse
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any

from _gatelib import git_text as git
from _gatelib import run_gate

EMPTY_TREE = "4b825dc642cb6eb9a060e54bf8d69288fbee4904"
TOOL_NAME = "Committed whitespace"




def resolve_commit(root: Path, revision: str) -> str:
    completed = git(root, "rev-parse", "--verify", f"{revision}^{{commit}}")
    if completed.returncode != 0:
        detail = completed.stderr.strip() or completed.stdout.strip()
        raise RuntimeError(f"Cannot resolve commit {revision!r}: {detail or 'unknown Git error'}")
    return completed.stdout.strip()


def resolve_committed_scope(
    root: Path, head_revision: str, base_revision: str | None
) -> tuple[str, str, str]:
    head = resolve_commit(root, head_revision)
    if base_revision:
        base = resolve_commit(root, base_revision)
        completed = git(root, "merge-base", base, head)
        if completed.returncode != 0 or not completed.stdout.strip():
            detail = completed.stderr.strip() or completed.stdout.strip()
            raise RuntimeError(
                "Cannot compute merge base for "
                f"{base_revision!r} and {head_revision!r}: {detail or 'unknown Git error'}"
            )
        return completed.stdout.strip(), head, "merge-base"

    parent = git(root, "rev-parse", "--verify", f"{head}^1")
    if parent.returncode == 0 and parent.stdout.strip():
        return parent.stdout.strip(), head, "first-parent"
    return EMPTY_TREE, head, "empty-tree"


def run_check(
    root: Path,
    mode: str,
    head_revision: str = "HEAD",
    base_revision: str | None = None,
) -> dict[str, Any]:
    if mode == "committed":
        base, head, basis = resolve_committed_scope(root, head_revision, base_revision)
        completed = git(root, "diff", "--check", base, head, "--")
        detail = (completed.stdout + completed.stderr).strip()
        return {
            "schema_version": 1,
            "tool": TOOL_NAME,
            "mode": mode,
            "status": "pass" if completed.returncode == 0 else "fail",
            "basis": basis,
            "base": base,
            "head": head,
            "range": f"{base}..{head}",
            "findings": detail.splitlines() if detail else [],
        }

    if base_revision is not None:
        raise ValueError("--base is valid only in committed mode.")
    if head_revision != "HEAD":
        raise ValueError("--head is valid only in committed mode.")

    unstaged = git(root, "diff", "--check", "--")
    staged = git(root, "diff", "--cached", "--check")
    findings: list[str] = []
    if unstaged.returncode != 0:
        findings.extend(
            f"unstaged: {line}"
            for line in (unstaged.stdout + unstaged.stderr).strip().splitlines()
            if line
        )
    if staged.returncode != 0:
        findings.extend(
            f"staged: {line}"
            for line in (staged.stdout + staged.stderr).strip().splitlines()
            if line
        )
    return {
        "schema_version": 1,
        "tool": TOOL_NAME,
        "mode": mode,
        "status": "pass" if not findings else "fail",
        "basis": "working-tree",
        "base": resolve_commit(root, "HEAD"),
        "head": None,
        "range": "working-tree",
        "findings": findings,
    }


def markdown_report(report: dict[str, Any]) -> str:
    lines = [
        "## Committed whitespace" if report["mode"] == "committed" else "## Working-tree whitespace",
        "",
        f"- **Status:** {str(report['status']).upper()}",
        f"- **Mode:** `{report['mode']}`",
        f"- **Basis:** `{report['basis']}`",
    ]
    if report["mode"] == "committed":
        lines.extend(
            [
                f"- **Base:** `{report['base']}`",
                f"- **Head:** `{report['head']}`",
                f"- **Inspected range:** `{report['range']}`",
            ]
        )
    else:
        lines.append("- **Inspected range:** staged and unstaged changes against `HEAD`")
    findings = report["findings"]
    lines.extend(["", f"**Findings:** {len(findings)}"])
    if findings:
        lines.extend(["", "```text", *findings, "```"])
    return "\n".join(lines) + "\n"




def init_repo(root: Path) -> None:
    git(root, "init", "-b", "main", check=True)
    git(root, "config", "user.name", "Whitespace Self-Test", check=True)
    git(root, "config", "user.email", "whitespace@example.invalid", check=True)


def commit_file(root: Path, text: str, message: str) -> str:
    (root / "fixture.txt").write_text(text, encoding="utf-8", newline="\n")
    git(root, "add", "fixture.txt", check=True)
    git(root, "commit", "-m", message, check=True)
    return resolve_commit(root, "HEAD")


def run_self_test() -> int:
    failures: list[str] = []

    with tempfile.TemporaryDirectory(prefix="committed-whitespace-root-bad-") as temporary:
        root = Path(temporary)
        init_repo(root)
        commit_file(root, "bad  \n", "root bad")
        report = run_check(root, "committed")
        if report["status"] != "fail" or report["basis"] != "empty-tree":
            failures.append("root commit with trailing whitespace was not rejected against empty tree")

    with tempfile.TemporaryDirectory(prefix="committed-whitespace-commit-bad-") as temporary:
        root = Path(temporary)
        init_repo(root)
        first = commit_file(root, "clean\n", "clean root")
        second = commit_file(root, "bad  \n", "bad second")
        report = run_check(root, "committed")
        if report["status"] != "fail" or report["base"] != first or report["head"] != second:
            failures.append("committed trailing whitespace was not rejected from a clean checkout")

    with tempfile.TemporaryDirectory(prefix="committed-whitespace-clean-") as temporary:
        root = Path(temporary)
        init_repo(root)
        first = commit_file(root, "clean\n", "clean root")
        second = commit_file(root, "still clean\n", "clean second")
        report = run_check(root, "committed")
        if report["status"] != "pass" or report["base"] != first or report["head"] != second:
            failures.append("valid clean commit did not pass")

        (root / "fixture.txt").write_text("unstaged  \n", encoding="utf-8", newline="\n")
        committed = run_check(root, "committed")
        local = run_check(root, "working-tree")
        if committed["status"] != "pass":
            failures.append("committed mode was polluted by an unstaged working-tree defect")
        if local["status"] != "fail":
            failures.append("working-tree mode did not reject an unstaged defect")

        git(root, "add", "fixture.txt", check=True)
        staged = run_check(root, "working-tree")
        if staged["status"] != "fail" or not any(
            str(item).startswith("staged:") for item in staged["findings"]
        ):
            failures.append("working-tree mode did not reject a staged defect")

    with tempfile.TemporaryDirectory(prefix="committed-whitespace-explicit-base-") as temporary:
        root = Path(temporary)
        init_repo(root)
        first = commit_file(root, "clean\n", "clean root")
        second = commit_file(root, "cleaner\n", "clean second")
        report = run_check(root, "committed", head_revision=second, base_revision=first)
        if (
            report["status"] != "pass"
            or report["basis"] != "merge-base"
            or report["base"] != first
            or report["head"] != second
        ):
            failures.append("explicit-base committed scope did not resolve through merge-base")

    if failures:
        for message in failures:
            print(f"[FAIL] {message}")
        print(f"SELF-TEST FAIL: {len(failures)} failure(s).")
        return 1
    print(
        "SELF-TEST PASS: root-commit, committed-defect, clean-commit, explicit-base, "
        "staged/unstaged, and committed-vs-working-tree separation fixtures passed."
    )
    return 0


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path.cwd())
    parser.add_argument(
        "--mode",
        choices=("committed", "working-tree"),
        default="committed",
        help="Inspect committed candidate changes or local staged/unstaged changes.",
    )
    parser.add_argument(
        "--head",
        default="HEAD",
        help="Committed-mode candidate revision (default: HEAD).",
    )
    parser.add_argument(
        "--base",
        help="Optional committed-mode base revision; the merge base with --head is inspected.",
    )
    parser.add_argument("--output", type=Path, help="Write deterministic JSON evidence.")
    parser.add_argument("--summary", type=Path, help="Write Markdown evidence.")
    parser.add_argument("--self-test", action="store_true")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    options = parse_args(sys.argv[1:] if argv is None else argv)
    return run_gate(
        options,
        build=lambda: run_check(
            options.root,
            options.mode,
            head_revision=options.head,
            base_revision=options.base,
        ),
        markdown=markdown_report,
        errors=(OSError, subprocess.SubprocessError, RuntimeError, ValueError),
        self_test=run_self_test,
        self_test_error_prefix="SELF-TEST ERROR",
    )

if __name__ == "__main__":
    raise SystemExit(main())
