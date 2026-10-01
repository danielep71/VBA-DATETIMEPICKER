#!/usr/bin/env python3
"""Private standard-library primitives shared by focused repository gates.

The canonical portable checker, ``check_repo.py``, intentionally does not import
this module: that file remains a self-contained distributable artifact.
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from collections.abc import Callable
from pathlib import Path
from typing import Any


def git_bytes(
    root: Path, *args: str, check: bool = False
) -> subprocess.CompletedProcess[bytes]:
    return subprocess.run(
        ["git", "-C", str(root), *args],
        check=check,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )


def git_text(
    root: Path, *args: str, check: bool = False
) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        ["git", "-C", str(root), *args],
        check=check,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )


def tracked_files(root: Path) -> set[str]:
    completed = git_bytes(root, "ls-files", "-z")
    if completed.returncode:
        raise RuntimeError(
            completed.stderr.decode("utf-8", errors="replace").strip()
        )
    return {
        item.decode("utf-8", errors="surrogateescape")
        for item in completed.stdout.split(b"\0")
        if item
    }


def write_text(path: Path | None, text: str) -> None:
    if path is None:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\n")


def write_json(path: Path | None, value: object) -> None:
    write_text(path, json.dumps(value, indent=2, sort_keys=True) + "\n")


def parse_report_args(
    argv: list[str], *, description: str | None = None
) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument("--root", type=Path, default=Path.cwd())
    parser.add_argument("--output", type=Path)
    parser.add_argument("--summary", type=Path)
    parser.add_argument("--self-test", action="store_true")
    return parser.parse_args(argv)


def run_gate(
    options: argparse.Namespace,
    *,
    build: Callable[[], dict[str, Any]],
    markdown: Callable[[dict[str, Any]], str],
    errors: tuple[type[Exception], ...],
    self_test: Callable[[], int] | None = None,
    self_test_error_prefix: str = "ERROR",
) -> int:
    """Run the shared console, report and exit contract of one focused gate.

    ``build`` produces the gate's report mapping and ``markdown`` renders it.
    ``errors`` is the gate's own operational-exception tuple: it is never
    widened here, so a programming error still propagates as a traceback
    instead of being reported as exit code 2.

    Self-test dispatch keeps each gate's existing diagnostic prefix. Gates that
    historically evaluated ``--self-test`` outside their operational handler
    reported failures as ``SELF-TEST ERROR``; gates that evaluated it inside
    reported ``ERROR``. Callers select that wording with
    ``self_test_error_prefix``.

    Gates whose output contract differs -- ``check_release.py`` (atomic writes
    and a console rendering distinct from its Markdown summary),
    ``test_workflow_validation.py`` (text-only report, no JSON output) and
    ``check_repo.py`` (a self-contained distributable that must not import this
    module) -- keep their own ``main``. The template's complete focused-tool
    consumer/exclusion registry lives in ``checker_development.py``; the
    canonical checker is independently excluded by its single-file contract.
    """
    if self_test is not None and options.self_test:
        try:
            return self_test()
        except errors as error:
            print(f"{self_test_error_prefix}: {error}", file=sys.stderr)
            return 2
    try:
        report = build()
        summary = markdown(report)
        write_json(options.output, report)
        write_text(options.summary, summary)
        print(summary.rstrip())
        return 0 if report["status"] == "pass" else 1
    except errors as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2
