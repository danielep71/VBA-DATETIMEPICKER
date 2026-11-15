"""Guard the trigger filters of the static-checks and CodeQL workflows.

Pull requests to any base run both workflows, so a pull request stacked on a
task branch is checked too. Pushes run only where they land: ``main`` and the
release branches. ``release-check.yml`` selects the static-checks *push* run on
``main`` for a release candidate, so dropping ``main`` from that filter would
break release validation without any other failure.

The workflows are read as text with the standard library, because the tooling
tests must run without third-party packages.
"""
import re
import unittest
from pathlib import Path

WORKFLOWS = Path(__file__).resolve().parent.parent / ".github" / "workflows"
GUARDED = ("static-checks.yml", "codeql.yml")
REQUIRED_PUSH_BRANCHES = {"main", "release/**"}


def trigger_block(text: str) -> list[str]:
    """Return the lines of the top-level ``on:`` mapping, comments removed."""
    lines = text.splitlines()
    start = next(i for i, line in enumerate(lines) if re.match(r"^on:\s*$", line))
    block: list[str] = []
    for line in lines[start + 1:]:
        if line and not line[0].isspace() and not line.startswith("#"):
            break
        stripped = line.split(" #", 1)[0].rstrip()
        if stripped.strip() and not stripped.lstrip().startswith("#"):
            block.append(stripped)
    return block


def event_body(block: list[str], event: str) -> list[str] | None:
    """Return the lines nested under one event, or None if the event is absent."""
    for index, line in enumerate(block):
        if re.match(rf"^  {re.escape(event)}:\s*(.*)$", line):
            body: list[str] = []
            for nested in block[index + 1:]:
                if re.match(r"^  \S", nested):
                    break
                body.append(nested)
            inline = re.match(rf"^  {re.escape(event)}:\s*(.*)$", line)
            if inline and inline.group(1):
                body.insert(0, "    " + inline.group(1))
            return body
    return None


def branches(body: list[str]) -> set[str]:
    """Return the branch patterns of an event, in flow or block style."""
    found: set[str] = set()
    for index, line in enumerate(body):
        match = re.match(r"^\s+branches:\s*(.*)$", line)
        if not match:
            continue
        flow = match.group(1).strip()
        if flow.startswith("["):
            found.update(item.strip().strip("'\"") for item in flow.strip("[]").split(",") if item.strip())
            continue
        for item in body[index + 1:]:
            entry = re.match(r"^\s+-\s*(.+)$", item)
            if not entry:
                break
            found.add(entry.group(1).strip().strip("'\""))
    return found


class WorkflowTriggerTests(unittest.TestCase):
    def read(self, name: str) -> str:
        return (WORKFLOWS / name).read_text(encoding="utf-8")

    def test_pull_requests_to_any_base_are_checked(self) -> None:
        for name in GUARDED:
            with self.subTest(workflow=name):
                body = event_body(trigger_block(self.read(name)), "pull_request")
                self.assertIsNotNone(body, "the workflow must run on pull_request")
                assert body is not None
                self.assertFalse(any(re.match(r"^\s+(branches|branches-ignore):", line) for line in body),
                                 "pull_request must have no branch filter, so stacked pull requests run it")

    def test_pushes_run_where_they_land(self) -> None:
        for name in GUARDED:
            with self.subTest(workflow=name):
                body = event_body(trigger_block(self.read(name)), "push")
                self.assertIsNotNone(body, "the workflow must run on push")
                assert body is not None
                self.assertEqual(branches(body), REQUIRED_PUSH_BRANCHES,
                                 "push must run on main and the release branches only")

    def test_static_checks_jobs_are_never_skipped_by_event(self) -> None:
        # A skipped run of the gating workflow on a commit reads as passing, so
        # duplicates must be avoided by trigger filters, never by a job-level if.
        # CodeQL's push and pull-request jobs are complementary by design and
        # are not covered by this rule.
        jobs = self.read("static-checks.yml").split("\njobs:\n", 1)[1]
        self.assertIsNone(re.search(r"^    if:.*github\.event_name", jobs, re.MULTILINE))

    def test_release_check_reads_the_push_run_on_main(self) -> None:
        text = self.read("release-check.yml")
        self.assertIn("static-checks.yml", text)
        self.assertIn('r["event"] == "push"', text)
        self.assertIn('r["head_branch"] == "main"', text)


class TriggerParserTests(unittest.TestCase):
    def test_parser_detects_regressions(self) -> None:
        filtered_pr = 'on:\n  pull_request:\n    branches: [main]\n  push:\n    branches: [main, "release/**"]\njobs:\n'
        body = event_body(trigger_block(filtered_pr), "pull_request")
        assert body is not None
        self.assertTrue(any("branches" in line for line in body))
        no_main = 'on:\n  push:\n    branches:\n      - "release/**"\n  pull_request:\njobs:\n'
        push = event_body(trigger_block(no_main), "push")
        assert push is not None
        self.assertEqual(branches(push), {"release/**"})
        self.assertEqual(event_body(trigger_block(no_main), "pull_request"), [])
        self.assertIsNone(event_body(trigger_block(no_main), "schedule"))


if __name__ == "__main__":
    unittest.main()
