"""Guard how Certify-Excel.ps1 and Build-Addin.ps1 close Excel.

A workbook ``Close`` or an Excel ``Quit`` that fails must count against the
verdict (issue #196). Both scripts close Excel only through
``tools/ExcelCleanup.psm1``, record its result with ``Add-CleanupResult``
before the verdict, and pass only when no reason is listed.

PowerShell does not run in CI, so the scripts are read as text here.
``tools/Test-ExcelCleanup.ps1`` exercises the module itself on Windows.
"""
import re
import unittest
from pathlib import Path

TOOLS = Path(__file__).resolve().parent
SCRIPTS = ("Certify-Excel.ps1", "Build-Addin.ps1")
MODULE = TOOLS / "ExcelCleanup.psm1"
IMPORT = 'Import-Module (Join-Path $PSScriptRoot "ExcelCleanup.psm1") -Force'
VERDICT = "if ($summary.reasons.Count -eq 0) {"


def code_lines(text: str) -> list[str]:
    """Return the script's lines without the comment-based help and comments."""
    text = re.sub(r"<#.*?#>", "", text, flags=re.S)
    return [line for line in text.splitlines() if not line.lstrip().startswith("#")]


class ScriptTests(unittest.TestCase):
    def setUp(self) -> None:
        self.scripts = {name: (TOOLS / name).read_text(encoding="ascii") for name in SCRIPTS}

    def test_scripts_import_the_cleanup_module(self) -> None:
        for name, text in self.scripts.items():
            with self.subTest(script=name):
                self.assertIn(IMPORT, text)

    def test_scripts_do_not_close_excel_themselves(self) -> None:
        for name, text in self.scripts.items():
            code = "\n".join(code_lines(text))
            with self.subTest(script=name):
                self.assertNotRegex(code, r"\.Quit\(")
                self.assertNotRegex(code, r"\.Close\(")
                self.assertNotIn("ReleaseComObject", code)
                self.assertIn("Close-ExcelInstance", code)

    def test_cleanup_is_recorded_before_the_verdict(self) -> None:
        for name, text in self.scripts.items():
            code = code_lines(text)
            with self.subTest(script=name):
                record = [i for i, line in enumerate(code) if line.startswith("Add-CleanupResult -Summary $summary")]
                verdict = [i for i, line in enumerate(code) if line == VERDICT]
                self.assertEqual(len(record), 1, "Add-CleanupResult must run once, at script level")
                self.assertEqual(len(verdict), 1, "one verdict")
                self.assertLess(record[0], verdict[0])
                self.assertEqual(code[verdict[0] + 1].strip(), '$summary.result = "PASS"')

    def test_pass_is_set_only_by_the_verdict(self) -> None:
        for name, text in self.scripts.items():
            code = code_lines(text)
            with self.subTest(script=name):
                assignments = [line.strip() for line in code if re.search(r"\$summary\.result\s*=", line)]
                self.assertEqual(assignments, ['$summary.result = "PASS"'])
                exits = [i for i, line in enumerate(code) if line.strip() == "exit 0"]
                self.assertEqual(len(exits), 1)
                self.assertEqual(code[exits[0] - 1].strip(), 'if ($summary.result -eq "PASS") {')

    def test_summary_has_a_cleanup_entry(self) -> None:
        for name, text in self.scripts.items():
            with self.subTest(script=name):
                self.assertIn('cleanup            = [ordered]@{ status = "NOT_RUN"; failures = @() }', text)
                self.assertIn('"Excel cleanup:   $($summary.cleanup.status)",', text)


class ModuleTests(unittest.TestCase):
    def setUp(self) -> None:
        self.text = MODULE.read_text(encoding="ascii")

    def test_module_exports_both_functions(self) -> None:
        self.assertIn("Export-ModuleMember -Function Close-ExcelInstance, Add-CleanupResult", self.text)

    def test_every_step_is_guarded(self) -> None:
        code = "\n".join(code_lines(self.text))
        for call in (r"\$book\.Close\(\$false\)", r"\$Excel\.Quit\(\)", r"ReleaseComObject\(\$book\)",
                     r"ReleaseComObject\(\$Excel\)"):
            with self.subTest(call=call):
                self.assertRegex(code, r"try \{\s*(?:if \([^\n]*\) \{\s*)?\[void\]" + r"[^\n]*" + call)

    def test_failures_become_reasons(self) -> None:
        self.assertIn('$Summary["reasons"] = @($Summary["reasons"]) + ("Excel cleanup: " + $failure)', self.text)

    def test_label_is_not_read_as_a_scoped_variable(self) -> None:
        # "$Label:" would be parsed as a variable in the scope named Label.
        self.assertNotIn('"$Label:', self.text)


if __name__ == "__main__":
    unittest.main()
