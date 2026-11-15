"""Guard the contract between the harness and tools/Certify-Excel.ps1 (#222).

The harness writes its summary to fixed cells of TST_DP_RESULTS, and
Certify-Excel.ps1 reads them back for the evidence. A package run must never
save the package and must prove its hash unchanged.

PowerShell and Excel do not run in CI, so both files are read as text here.
"""
import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
HARNESS = ROOT / "test" / "M_cDP_Test.bas"
SCRIPT = ROOT / "tools" / "Certify-Excel.ps1"


def code_lines(text: str) -> list[str]:
    """Return the script's lines without the comment-based help and comments."""
    text = re.sub(r"<#.*?#>", "", text, flags=re.S)
    return [line for line in text.splitlines() if not line.lstrip().startswith("#")]


def function_body(code: str, name: str) -> str:
    """Return the text of one PowerShell function, up to the next function."""
    match = re.search(rf"^function {re.escape(name)} \{{\n(.*?)^\}}\n", code, flags=re.S | re.M)
    if match is None:
        raise AssertionError(f"function {name} not found")
    return match.group(1)


class HarnessCellsTests(unittest.TestCase):
    def setUp(self) -> None:
        self.harness = HARNESS.read_text(encoding="ascii")
        self.code = "\n".join(code_lines(SCRIPT.read_text(encoding="ascii")))

    def test_summary_rows_match(self) -> None:
        written = dict(re.findall(
            r"mTST_DP_ResultSheet\.Cells\((\d+), TST_DP_COL_SUMMARY_VALUE\)\.Value = (\w+)", self.harness))
        self.assertEqual(written, {
            "5": "mTST_DP_RunCount", "6": "mTST_DP_PassCount", "7": "mTST_DP_FailCount",
            "8": "TST_DP_ResolveRunState", "9": "mTST_DP_SuitesDispatched", "10": "mTST_DP_SuitesCompleted"})
        read = dict((field, row) for field, row in re.findall(
            r"\$Record\.(\w+) = (?:ConvertTo-Count |\[string\])\$sheet\.Cells\.Item\((\d+), \$summaryValueColumn\)",
            self.code))
        self.assertEqual(read, {"run": "5", "passed": "6", "failed": "7", "state": "8",
                                "suites_dispatched": "9", "suites_completed": "10"})
        self.assertIn("Private Const TST_DP_COL_SUMMARY_VALUE  As Long = 11", self.harness)
        self.assertIn("$summaryValueColumn = 11", self.code)

    def test_summary_line_format_is_unchanged(self) -> None:
        self.assertIn('"State=" & TST_DP_ResolveRunState() & _', self.harness)
        self.assertIn('"; CleanupFailures=" & VBA.CStr(mTST_DP_CleanupFails)', self.harness)
        self.assertIn('TST_DP_RecordInfo "Harness", "Suites", _', self.harness)

    def test_incomplete_suites_fail_the_run(self) -> None:
        reasons = function_body(self.code, "Get-HarnessReasons")
        self.assertIn("$Record.suites_completed -ne $Record.suites_dispatched", reasons)
        self.assertIn("$Record.suites_dispatched -eq 0", reasons)
        self.assertIn("were not found", reasons)


class PackageRunTests(unittest.TestCase):
    def setUp(self) -> None:
        self.code = "\n".join(code_lines(SCRIPT.read_text(encoding="ascii")))
        start = self.code.index("    if ($packageMode) {\n        foreach ($key in $packs.Keys) {")
        end = self.code.index("    else {\n        $excel = Start-Excel", start)
        self.package_branch = self.code[start:end]

    def test_both_packs_run(self) -> None:
        self.assertIn('standard = "TST_DP_RunAll"', self.code)
        self.assertIn('ui_smoke = "TST_DP_RunAll_WithUISmoke"', self.code)
        self.assertIn("foreach ($key in $packs.Keys) {", self.package_branch)

    def test_package_is_opened_read_only_and_never_saved(self) -> None:
        self.assertIn("$excel.Workbooks.Open($packagePath, 0, $true)", self.package_branch)
        self.assertNotRegex(self.package_branch, r"\.Save(As|CopyAs)?\(")
        self.assertIn("Close-Excel -Excel $excel -Workbooks @($hostBook, $workbook)", self.package_branch)

    def test_hash_is_checked_after_each_pack(self) -> None:
        self.assertIn("$pack.sha256_after = Get-Sha256 -Path $packagePath", self.package_branch)
        self.assertIn("if ($pack.sha256_after -ne $summary.package.sha256) {", self.package_branch)
        verdict = self.code[self.code.index("Add-CleanupResult -Summary $summary"):]
        self.assertIn("if ($pack.sha256_after -ne $summary.package.sha256) {", verdict)

    def test_macros_are_enabled_for_this_instance_before_opening(self) -> None:
        enable = self.package_branch.index("Enable-PackageMacros -Excel $excel")
        self.assertLess(enable, self.package_branch.index("$excel.Workbooks.Open($packagePath, 0, $true)"))
        body = function_body(self.code, "Enable-PackageMacros")
        self.assertIn("$Excel.AutomationSecurity = $msoAutomationSecurityLow", body)
        self.assertNotIn("Set-ItemProperty", self.code)
        self.assertNotIn("New-ItemProperty", self.code)

    def test_host_records_what_release_evidence_needs(self) -> None:
        # check_release.py requires scaling, monitors, load_mode and
        # other_addins for each package; they come from the run, not by hand.
        for key in ("monitors", "scaling", "other_addins", "load_mode"):
            self.assertRegex(self.code, rf"\n        {key} +=")
        # Excel started by a script can still load COM add-ins (seen at 8a96dea),
        # so the load mode must not claim that none load.
        self.assertNotIn("no add-ins", self.code)

    def test_other_addin_names_stay_out_of_the_evidence(self) -> None:
        body = function_body(self.code, "Get-OtherAddIns")
        returns = re.findall(r"return (.+)", body)
        self.assertEqual(returns, ['"unknown"', '"unknown"', '"none"', '"$($names.Count) (names in session.log)"'])
        self.assertIn('Write-Session ("Other add-ins open: " + ($names -join "; "))', body)

    def test_workbook_qualifier_escapes_apostrophes(self) -> None:
        body = function_body(self.code, "Get-WorkbookQualifier")
        self.assertIn("""return "'" + ([string]$Workbook.Name).Replace("'", "''") + "'!\"""", body)
        self.assertEqual(self.code.count("$qualified = Get-WorkbookQualifier -Workbook $workbook"), 2)
        self.assertNotRegex(self.code, r"\$qualified = \"'\" \+")

    def test_package_runs_refuse_source_only_switches(self) -> None:
        self.assertIn("-WithUISmoke applies to a source run", self.code)


if __name__ == "__main__":
    unittest.main()
