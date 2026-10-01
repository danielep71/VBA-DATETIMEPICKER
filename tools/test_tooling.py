"""Offline negative fixtures for DateTimePicker-specific adaptations and evidence."""
import copy
import hashlib
import subprocess
import tempfile
import unittest
from pathlib import Path

import check_release
import check_source
import check_vba_conditionals as conditionals
import check_vba_jumps as jumps


class ConditionalTests(unittest.TestCase):
    def test_alternative_procedure_declarations(self):
        source = ('#If VBA7 Then\nPrivate Sub Test(ByVal p As LongPtr)\n'
                  '#Else\nPrivate Sub Test(ByVal p As Long)\n#End If\n'
                  'On Error GoTo Handler\nHandler:\nEnd Sub\n')
        sources, findings = conditionals.reachable_sources("test.bas", source)
        self.assertEqual(findings, [])
        for text in sources.values():
            self.assertEqual(jumps.analyze_component("test.bas", text), [])

    def test_windows_mac_branch_and_bad_reachable_declare(self):
        source = '#If Mac Then\nDeclare Sub MacOnly Lib "x" ()\n#Else\nDeclare Sub Bad Lib "x" ()\n#End If'
        findings = conditionals.analyze_component("test.bas", source)
        self.assertTrue(any(f.get("line") == 4 for f in findings))
        self.assertFalse(any(f.get("line") == 2 for f in findings))
        self.assertEqual(conditionals.analyze_component("test.bas", source.replace('Declare Sub Bad', 'Declare PtrSafe Sub Bad')), [])

    def test_cross_procedure_jump_still_fails(self):
        source = 'Sub A()\nGoTo Other\nEnd Sub\nSub B()\nOther:\nEnd Sub'
        self.assertTrue(jumps.analyze_component("test.bas", source))

    def test_ribbon_rejects_dtd(self):
        with self.assertRaises(ValueError):
            check_source.ribbon_callbacks(b'<!DOCTYPE x [<!ENTITY y "hello">]><x/>')
        self.assertEqual(check_source.ribbon_callbacks(b'<button onAction="Run"/>'), ["Run"])


class ReleaseTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        for args in (("init", "-q"), ("config", "user.name", "Fixture"),
                     ("config", "user.email", "fixture@example.invalid")):
            self.git(*args)
        (self.root / "VERSION").write_text("1.2.3\n")
        (self.root / "CHANGELOG.md").write_text("## [1.2.3] - 2026-10-01\n")
        self.git("add", ".")
        self.git("commit", "-qm", "fixture")
        self.sha = self.git("rev-parse", "HEAD").strip()
        self.git("update-ref", "refs/remotes/origin/main", self.sha)
        self.git("tag", "-a", "v1.2.3", "-m", "fixture")
        self.record = {"schema_version": 1, "candidate_sha": self.sha,
                       "version": "1.2.3", "tag": "v1.2.3", "assets": [],
                       "manual": {name: {"status": "PASS", "detail": "Synthetic fixture"}
                                  for name in check_release.MANUAL}}
        for host, filename in (("xlam", "DATETIMEPICKER.v1.2.3.xlam"),
                               ("xlsm", "DATETIMEPICKER-demo-v1.2.3.xlsm")):
            data = b"Synthetic bytes; not an Office certification"
            (self.root / filename).write_bytes(data)
            pack = {"state": "PASS", "run": 10, "passed": 10, "failed": 0,
                    "cleanup_failures": 0, "expected_run": 10, "suites": 2, "expected_suites": 2}
            self.record["assets"].append({"host": host, "filename": filename,
                "size": len(data), "sha256": hashlib.sha256(data).hexdigest(),
                "candidate_sha": self.sha, "compile": "PASS", "standard": pack,
                "ui_smoke": copy.deepcopy(pack), "environment": {k: "fixture" for k in
                ("excel", "windows", "bitness", "scaling", "monitors", "load_mode", "other_addins")}})
        self.static = {"status": "pass", "candidate_sha": self.sha, "dirty": False,
                       "mode": "committed", "checks": [{"name": name, "exit_code": 0}
                                                         for name in check_release.STATIC_CHECKS]}

    def git(self, *args):
        return subprocess.run(["git", "-C", str(self.root), *args], check=True,
                              capture_output=True, text=True).stdout

    def validate(self):
        return check_release.validate(self.root, self.record, self.root, self.static, True)

    def test_valid_and_accepted_gap_are_distinct(self):
        self.assertEqual(self.validate()["status"], "pass")
        item = self.record["manual"]["grid-ownership"]
        item.update(status="NOT RUN", detail="Explicit gap")
        with self.assertRaisesRegex(ValueError, "accepted"):
            self.validate()
        item["accepted_by"] = "Maintainer"
        self.assertEqual(self.validate()["status"], "pass_with_limitations")
        item["status"] = "FAIL"
        with self.assertRaisesRegex(ValueError, "failure"):
            self.validate()

    def test_modified_asset_fails(self):
        path = self.root / self.record["assets"][0]["filename"]
        path.write_bytes(b"x" * path.stat().st_size)
        with self.assertRaisesRegex(ValueError, "SHA-256"):
            self.validate()

    def test_preparation_does_not_invent_excel_results(self):
        record = check_release.prepare(self.root, self.sha, self.root)
        self.assertEqual(record["assets"][0]["compile"], "NOT RUN")
        self.assertEqual(record["assets"][1]["ui_smoke"]["state"], "NOT RUN")
        with self.assertRaisesRegex(ValueError, "compile"):
            check_release.validate(self.root, record, self.root, self.static)

    def test_missing_static_gate_fails(self):
        self.static["checks"].pop()
        with self.assertRaisesRegex(ValueError, "incomplete"):
            self.validate()

    def test_incomplete_regression_fails(self):
        self.record["assets"][1]["ui_smoke"]["expected_run"] += 1
        with self.assertRaisesRegex(ValueError, "incomplete"):
            self.validate()

    def test_stale_or_dirty_static_report_fails(self):
        for key, value in (("candidate_sha", "0" * 40), ("dirty", True), ("mode", "working-tree")):
            original = self.static[key]
            self.static[key] = value
            with self.assertRaisesRegex(ValueError, "Static report"):
                self.validate()
            self.static[key] = original

    def test_wrong_tag_fails(self):
        self.git("tag", "-d", "v1.2.3")
        self.git("tag", "v1.2.3")
        with self.assertRaisesRegex(ValueError, "Annotated"):
            self.validate()

    def test_unmerged_candidate_fails_even_when_tagged(self):
        self.git("commit", "--allow-empty", "-qm", "unmerged candidate")
        unmerged = self.git("rev-parse", "HEAD").strip()
        self.git("tag", "-fa", "v1.2.3", "-m", "unmerged candidate")
        self.record["candidate_sha"] = unmerged
        self.static["candidate_sha"] = unmerged
        for asset in self.record["assets"]:
            asset["candidate_sha"] = unmerged
        with self.assertRaisesRegex(ValueError, "reachable from fetched origin/main"):
            self.validate()


if __name__ == "__main__":
    unittest.main()
