"""Fixtures for tools/assemble_release.py: complete, incomplete and inconsistent input (#222).

The evidence folders are synthetic copies of what tools/Certify-Excel.ps1 and
the build scripts write; the packages are minimal Office-shaped archives.
"""
import argparse
import copy
import hashlib
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from typing import Any

import assemble_release
import check_release
from test_tooling import write_clean_package

HARNESS = (
    "Private Sub TST_DP_RunAllInternal()\n"
    '        TST_DP_RunSuiteSafe "Environment"\n'
    '        TST_DP_RunSuiteSafe "Settings"\n'
    "        If IncludeUISmoke Then\n"
    '            TST_DP_RunSuiteSafe "UISmoke"\n'
    "        End If\n"
    "End Sub\n"
)
CHANGELOG = (
    "# Changelog\n\n"
    "## [Unreleased]\n\n"
    '<a id="v123"></a>\n\n'
    "## [1.2.3] - 2026-11-24\n\n"
    "### Fixed\n\n"
    "- A fix described in [the manual](docs/USER_MANUAL.md) and [#5](https://example.invalid/5).\n\n"
    "---\n\n"
    "## [1.2.2] - 2026-10-01\n\n"
    "- Older.\n"
)
TOTALS = {"standard": 20, "ui_smoke": 23}
SUITES = {"standard": 2, "ui_smoke": 3}


def harness(pack: str, *, run: int | None = None, suites: int | None = None) -> dict[str, Any]:
    total = TOTALS[pack] if run is None else run
    dispatched = SUITES[pack] if suites is None else suites
    return {"status": "RAN", "error": None, "state": "PASS", "run": total, "passed": total, "failed": 0,
            "cleanup_failures": 0, "suites_dispatched": dispatched, "suites_completed": dispatched,
            "summary_line": f"State=PASS; Run={total}; Passed={total}; Failed=0; CleanupFailures=0"}


def host() -> dict[str, Any]:
    return {"os": "Windows 11 (fixture)", "excel_build": "Version 2611 (build 1)", "office_bitness": "64-bit",
            "scaling": "100%", "monitors": "1", "other_addins": "none",
            "load_mode": "xlam opened read-only by tools/Certify-Excel.ps1"}


class AssembleTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        base = Path(self.temp.name)
        self.root = base / "repo"
        self.root.mkdir()
        for args in (("init", "-q"), ("config", "user.name", "Fixture"),
                     ("config", "user.email", "fixture@example.invalid")):
            self.git(*args)
        (self.root / "VERSION").write_text("1.2.2\n")
        self.git("add", ".")
        self.git("commit", "-qm", "previous release")
        self.git("tag", "-a", "v1.2.2", "-m", "previous")
        (self.root / "VERSION").write_text("1.2.3\n")
        (self.root / "CHANGELOG.md").write_text(CHANGELOG)
        (self.root / "test").mkdir()
        (self.root / "test" / "M_cDP_Test.bas").write_text(HARNESS)
        self.git("add", ".")
        self.git("commit", "-qm", "candidate")
        self.sha = self.git("rev-parse", "HEAD").strip()
        self.git("update-ref", "refs/remotes/origin/main", self.sha)

        self.assets = base / "release-assets"
        self.assets.mkdir()
        self.hashes: dict[str, tuple[str, int]] = {}
        for host_name, name in (("xlam", "DATETIMEPICKER.v1.2.3.xlam"), ("xlsm", "DATETIMEPICKER-demo-v1.2.3.xlsm")):
            write_clean_package(self.assets / name)
            data = (self.assets / name).read_bytes()
            self.hashes[host_name] = (hashlib.sha256(data).hexdigest(), len(data))
        (self.assets / "DATETIMEPICKER-user-manual-v1.2.3.pdf").write_bytes(b"%PDF-1.7 fixture")

        self.evidence = base / "evidence"
        self.summaries: dict[str, dict[str, Any]] = {}
        for pack, entry in assemble_release.PACKS.items():
            self.summaries[f"source-{pack}"] = self.common(assemble_release.CERTIFY) | {
                "mode": "source", "entry_point": entry, "harness": harness(pack), "host": host()}
        for host_name in ("xlam", "xlsm"):
            digest, size = self.hashes[host_name]
            self.summaries[f"{host_name}-build"] = self.common(assemble_release.BUILDER) | {
                "kind": assemble_release.KINDS[host_name],
                "package": {"file": "built", "bytes": size, "sha256": digest, "sha256_after_reopen": digest},
                "metadata": {"status": "PASS"}}
            self.summaries[f"{host_name}-run"] = self.common(assemble_release.CERTIFY) | {
                "mode": "package", "host": host(),
                "package": {"file": "built", "host": host_name, "bytes": size, "sha256": digest},
                "packs": {pack: {"entry_point": entry, "compile": {"status": "PASS"}, "harness": harness(pack),
                                 "sha256_after": digest} for pack, entry in assemble_release.PACKS.items()}}
        self.manual = {name: {"status": "PASS", "detail": "Fixture check"} for name in check_release.MANUAL}
        self.static = {"schema_version": 1, "status": "pass", "candidate_sha": self.sha, "dirty": False,
                       "mode": "committed",
                       "checks": [{"name": name, "exit_code": 0} for name in check_release.STATIC_CHECKS]}

    def common(self, tool: str) -> dict[str, Any]:
        return {"tool": tool, "candidate_sha": self.sha, "working_tree_clean": True, "result": "PASS",
                "reasons": []}

    def git(self, *args: str) -> str:
        return subprocess.run(["git", "-C", str(self.root), *args], check=True,
                              capture_output=True, text=True).stdout

    def args(self) -> argparse.Namespace:
        folders = {}
        for name, summary in self.summaries.items():
            folder = self.evidence / name
            folder.mkdir(parents=True, exist_ok=True)
            (folder / "summary.json").write_text(json.dumps(summary), encoding="utf-8")
            folders[name] = folder
        manual = self.evidence / "manual-checks.json"
        manual.write_text(json.dumps(self.manual))
        static = self.evidence / "static-checks.json"
        static.write_text(json.dumps(self.static))
        return argparse.Namespace(
            root=self.root, candidate_sha=self.sha, assets_dir=self.assets, static_report=static,
            source_standard=folders["source-standard"], source_ui_smoke=folders["source-ui_smoke"],
            addin_build=folders["xlam-build"], addin_run=folders["xlam-run"],
            demo_build=folders["xlsm-build"], demo_run=folders["xlsm-run"], manual=manual,
            title="Fixture release", refs=[9, 7, 9], evidence_issue=122,
            output_dir=self.evidence / "out")

    def assemble(self) -> dict[str, str]:
        return assemble_release.assemble(self.args())

    # Complete input -----------------------------------------------------------

    def test_complete_input_gives_a_valid_record(self) -> None:
        files = self.assemble()
        record = json.loads(files["release-evidence.json"])
        self.assertEqual(json.loads(files["release-check.json"])["status"], "pass")
        self.assertEqual(check_release.validate(self.root, record, self.assets, self.static)["status"], "pass")
        xlam = next(asset for asset in record["assets"] if asset["host"] == "xlam")
        self.assertEqual(xlam["standard"], {"state": "PASS", "run": 20, "passed": 20, "failed": 0,
                                            "cleanup_failures": 0, "expected_run": 20, "suites": 2,
                                            "expected_suites": 2})
        self.assertEqual(xlam["ui_smoke"]["expected_suites"], 3)
        self.assertEqual(xlam["environment"]["other_addins"], "none")
        self.assertEqual(record["documents"][0]["filename"], "DATETIMEPICKER-user-manual-v1.2.3.pdf")

    def test_tag_message_follows_the_template(self) -> None:
        self.assertEqual(self.assemble()["tag-message.txt"], (
            "v1.2.3 - Fixture release\n\n"
            f"Certified at {self.sha}.\n"
            "Embedded .xlsm: standard 20/20, UI smoke 23/23, Failed=0, CleanupFailures=0.\n"
            "Packaged .xlam: standard 20/20, UI smoke 23/23, Failed=0, CleanupFailures=0.\n"
            "Refs #7, #9.\n"
            "Deviations: none.\n"
            "Evidence: #122.\n"))

    def test_release_notes_carry_changes_hashes_and_results(self) -> None:
        notes = self.assemble()["release-notes.md"]
        self.assertIn("# v1.2.3 — Fixture release", notes)
        self.assertIn("Released 2026-11-24.", notes)
        self.assertIn(f"[the manual]({assemble_release.REPOSITORY_URL}/blob/v1.2.3/docs/USER_MANUAL.md)", notes)
        self.assertIn("[#5](https://example.invalid/5)", notes)
        self.assertNotIn("Older.", notes)
        self.assertNotIn("Unreleased", notes)
        for host_name in ("xlam", "xlsm"):
            self.assertIn(f"`{self.hashes[host_name][0]}`", notes)
        self.assertIn("| Add-in `.xlam` | With UI smoke | 23/23 | 3/3 | 0 | 0 |", notes)
        self.assertIn(f"{assemble_release.REPOSITORY_URL}/compare/v1.2.2...v1.2.3", notes)

    def test_accepted_gap_is_a_deviation(self) -> None:
        self.manual["grid-ownership"] = {"status": "NOT RUN", "detail": "Two workbooks not available",
                                         "accepted_by": "Maintainer"}
        files = self.assemble()
        self.assertEqual(json.loads(files["release-check.json"])["status"], "pass_with_limitations")
        self.assertIn("Deviations: grid-ownership not run (Two workbooks not available; accepted by Maintainer).",
                      files["tag-message.txt"])

    # Incomplete input ---------------------------------------------------------

    def test_missing_evidence_folder_fails(self) -> None:
        args = self.args()
        args.demo_run = self.evidence / "absent"
        with self.assertRaisesRegex(ValueError, "xlsm package run: .* not found"):
            assemble_release.assemble(args)

    def test_missing_pdf_fails(self) -> None:
        (self.assets / "DATETIMEPICKER-user-manual-v1.2.3.pdf").unlink()
        with self.assertRaisesRegex(ValueError, "user-manual"):
            self.assemble()

    def test_incomplete_suites_fail(self) -> None:
        packs = self.summaries["xlam-run"]["packs"]
        packs["ui_smoke"]["harness"]["suites_completed"] = 2
        with self.assertRaisesRegex(ValueError, "2 of 3 suites completed"):
            self.assemble()

    def test_unanswered_manual_check_fails(self) -> None:
        del self.manual["clock-lifecycle"]
        with self.assertRaisesRegex(ValueError, "expected exactly"):
            self.assemble()
        self.manual["clock-lifecycle"] = {"status": "NOT RUN", "detail": "Not done"}
        with self.assertRaisesRegex(ValueError, "accepted"):
            self.assemble()

    def test_unknown_environment_fails(self) -> None:
        self.summaries["xlsm-run"]["host"]["excel_build"] = "unknown"
        with self.assertRaisesRegex(ValueError, "excel is unknown"):
            self.assemble()

    # Inconsistent input -------------------------------------------------------

    def test_package_totals_must_equal_the_source_run(self) -> None:
        self.summaries["source-standard"]["harness"] = harness("standard", run=21)
        with self.assertRaisesRegex(ValueError, "incomplete"):
            self.assemble()

    def test_suites_must_match_the_harness_source(self) -> None:
        self.summaries["source-ui_smoke"]["harness"] = harness("ui_smoke", suites=4)
        with self.assertRaisesRegex(ValueError, "4 suites dispatched, the harness has 3"):
            self.assemble()

    def test_package_suites_must_match_the_harness_source(self) -> None:
        self.summaries["xlsm-run"]["packs"]["ui_smoke"]["harness"] = harness("ui_smoke", suites=4)
        with self.assertRaisesRegex(ValueError, "incomplete"):
            self.assemble()

    def test_package_changed_after_reopening_fails(self) -> None:
        self.summaries["xlam-build"]["package"]["sha256_after_reopen"] = "d" * 64
        with self.assertRaisesRegex(ValueError, "changed after reopening"):
            self.assemble()

    def test_another_candidate_fails(self) -> None:
        self.summaries["xlsm-build"]["candidate_sha"] = "0" * 40
        with self.assertRaisesRegex(ValueError, "xlsm build: candidate"):
            self.assemble()

    def test_dirty_or_failed_run_fails(self) -> None:
        original = copy.deepcopy(self.summaries["xlam-run"])
        self.summaries["xlam-run"]["working_tree_clean"] = False
        with self.assertRaisesRegex(ValueError, "not clean"):
            self.assemble()
        self.summaries["xlam-run"] = original
        self.summaries["xlam-run"]["result"] = "FAIL"
        self.summaries["xlam-run"]["reasons"] = ["Pack standard: compile FAIL"]
        with self.assertRaisesRegex(ValueError, "compile FAIL"):
            self.assemble()

    def test_tested_package_must_be_the_asset(self) -> None:
        self.summaries["xlam-run"]["package"]["sha256"] = "f" * 64
        with self.assertRaisesRegex(ValueError, "tested package is not"):
            self.assemble()
        self.summaries["xlam-run"]["package"]["sha256"] = self.hashes["xlam"][0]
        write_clean_package(self.assets / "DATETIMEPICKER.v1.2.3.xlam", creator="Someone")
        with self.assertRaisesRegex(ValueError, "built package is not"):
            self.assemble()

    def test_package_changed_during_a_pack_fails(self) -> None:
        self.summaries["xlsm-run"]["packs"]["standard"]["sha256_after"] = "e" * 64
        with self.assertRaisesRegex(ValueError, "changed during the pack"):
            self.assemble()

    def test_wrong_builder_fails(self) -> None:
        self.summaries["xlsm-build"]["kind"] = "addin"
        with self.assertRaisesRegex(ValueError, "xlsm build: built a 'addin' package"):
            self.assemble()
        self.summaries["xlsm-build"]["kind"] = "demo"
        self.summaries["xlsm-build"]["tool"] = assemble_release.CERTIFY
        with self.assertRaisesRegex(ValueError, "xlsm build: written by"):
            self.assemble()

    # Command line -------------------------------------------------------------

    def test_existing_output_is_never_overwritten(self) -> None:
        args = self.args()
        args.output_dir.mkdir()
        (args.output_dir / "release-evidence.json").write_text("{}")
        argv = ["--root", str(args.root), "--candidate-sha", self.sha, "--assets-dir", str(self.assets),
                "--static-report", str(args.static_report), "--source-standard", str(args.source_standard),
                "--source-ui-smoke", str(args.source_ui_smoke), "--addin-build", str(args.addin_build),
                "--addin-run", str(args.addin_run), "--demo-build", str(args.demo_build),
                "--demo-run", str(args.demo_run), "--manual", str(args.manual), "--title", "Fixture",
                "--evidence-issue", "122", "--output-dir", str(args.output_dir)]
        script = Path(assemble_release.__file__)
        result = subprocess.run([sys.executable, str(script), *argv], capture_output=True, text=True, check=False)
        self.assertEqual(result.returncode, 1)
        self.assertIn("is not empty", result.stdout)
        (args.output_dir / "release-evidence.json").unlink()
        result = subprocess.run([sys.executable, str(script), *argv], capture_output=True, text=True, check=False)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertEqual(sorted(p.name for p in args.output_dir.iterdir()), sorted(assemble_release.OUTPUTS))


if __name__ == "__main__":
    unittest.main()
