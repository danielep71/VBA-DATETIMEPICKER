"""Guard the demo workbook build: tools/Build-Addin.ps1 -Demo and demo/M_DP_DEMO.bas (#222).

The script calls one tracked VBA entry point by name and then checks the sheets
that entry point leaves. Both sides are read as text here, because PowerShell
and Excel do not run in CI.
"""
import re
import unittest
from pathlib import Path

import assemble_release

ROOT = Path(__file__).resolve().parent.parent
SCRIPT = ROOT / "tools" / "Build-Addin.ps1"
DEMO = ROOT / "demo" / "M_DP_DEMO.bas"
RUNTIME = ROOT / "src" / "modules" / "M_DatePicker.bas"
LAUNCHER = ROOT / "tools" / "Build-Demo.cmd"


def code_lines(text: str) -> list[str]:
    """Return the script's lines without the comment-based help and comments."""
    text = re.sub(r"<#.*?#>", "", text, flags=re.S)
    return [line for line in text.splitlines() if not line.lstrip().startswith("#")]


def vba_constant(text: str, name: str) -> str:
    match = re.search(rf'Private Const {name}\s+As String = "([^"]*)"', text)
    if match is None:
        raise AssertionError(f"{name} not found")
    return match.group(1)


def vba_procedure(text: str, header: str) -> str:
    start = text.index(header)
    end = re.search(r"^End (Sub|Function)$", text[start:], flags=re.M)
    if end is None:
        raise AssertionError(f"{header} has no end")
    return text[start:start + end.end()]


class DemoBuildTests(unittest.TestCase):
    def setUp(self) -> None:
        self.code = "\n".join(code_lines(SCRIPT.read_text(encoding="ascii")))
        self.demo = DEMO.read_text(encoding="ascii").replace("\r\n", "\n")
        self.runtime = RUNTIME.read_text(encoding="cp1252").replace("\r\n", "\n")
        start = self.code.index("    if ($Demo) {\n        $caption =")
        self.content = self.code[start:self.code.index("    $workbook.SaveAs(", start)]

    def test_entry_point_exists_and_never_raises(self) -> None:
        self.assertIn('$excel.Run($qualified + "DP_Demo_PrepareReleaseWorkbook", $caption)', self.content)
        body = vba_procedure(self.demo, "Public Function DP_Demo_PrepareReleaseWorkbook(")
        self.assertIn("On Error GoTo Failed", body)
        self.assertIn('DP_Demo_PrepareReleaseWorkbook = "OK"', body)
        self.assertIn('DP_Demo_PrepareReleaseWorkbook = "ERROR " & VBA.CStr(Err.Number) & ": " & Err.Description', body)
        self.assertNotIn("Err.Raise", body)
        self.assertIn('if ($result -ne "OK") {', self.content)

    def test_sheet_names_agree(self) -> None:
        self.assertEqual(vba_constant(self.demo, "DEMO_SHEET_NAME"), vba_constant(self.runtime, "DP_DEMO_SHEET_NAME"))
        self.assertIn(f'$demoSheetName = "{vba_constant(self.demo, "DEMO_SHEET_NAME")}"', self.code)
        self.assertIn(f'$startSheetName = "{vba_constant(self.demo, "START_SHEET_NAME")}"', self.code)

    def test_opening_state_is_checked(self) -> None:
        body = vba_procedure(self.demo, "Public Function DP_Demo_PrepareReleaseWorkbook(")
        self.assertIn("Wb.Worksheets(DEMO_SHEET_NAME).Visible = xlSheetVeryHidden", body)
        self.assertIn("$sheets[0].name -eq $startSheetName -and $sheets[0].visible -eq -1", self.content)
        self.assertIn("$sheets[1].name -eq $demoSheetName -and $sheets[1].visible -eq $xlSheetVeryHidden", self.content)
        start = vba_procedure(self.demo, "Public Sub DP_Demo_CreateStartSheet(")
        self.assertIn("WS.Move Before:=Wb.Worksheets(1)", start)

    def test_only_empty_default_sheets_are_removed(self) -> None:
        self.assertIn("[int]$excel.WorksheetFunction.CountA($sheet.Cells) -ne 0 -or [int]$sheet.Shapes.Count -ne 0",
                      self.content)
        self.assertIn('throw "The new workbook has a sheet that is not empty: $name"', self.content)

    def test_demo_is_saved_as_a_macro_enabled_workbook(self) -> None:
        self.assertIn("$xlOpenXMLWorkbookMacroEnabled = 52", self.code)
        self.assertIn("$workbook.SaveAs($packagePath, $(if ($Demo) { $xlOpenXMLWorkbookMacroEnabled } "
                      "else { $xlOpenXMLAddIn }))", self.code)
        self.assertIn('$extension = $(if ($Demo) { ".xlsm" } else { ".xlam" })', self.code)

    def test_reopen_enables_macros_for_this_instance_only(self) -> None:
        enable = self.code.index("$excel.AutomationSecurity = $msoAutomationSecurityLow")
        self.assertLess(enable, self.code.index("$addin = $excel.Workbooks.Open($packagePath, 0, $true)"))
        self.assertNotIn("Set-ItemProperty", self.code)

    def test_demo_questions_count_against_the_verdict(self) -> None:
        verdict = self.code[self.code.index("Add-CleanupResult -Summary $summary"):]
        for field in ("opens_on_start", "demo_sheet_works"):
            self.assertIn(f"$summary.ribbon_check.{field} = Read-YesNo", self.code)
            self.assertIn(f"if ($Demo -and -not $summary.ribbon_check.{field})", verdict)
        self.assertIn('if ($Demo -and $summary.demo_content.status -ne "PASS")', verdict)

    def test_summary_names_the_kind(self) -> None:
        self.assertIn('$kind = $(if ($Demo) { "demo" } else { "addin" })', self.code)
        self.assertIn("    kind               = $kind", self.code)

    def test_kinds_match_what_the_assembler_reads(self) -> None:
        self.assertEqual(assemble_release.BUILDER, "tools/Build-Addin.ps1")
        self.assertIn(f'tool               = "{assemble_release.BUILDER}"', self.code)
        self.assertEqual(assemble_release.KINDS, {"xlam": "addin", "xlsm": "demo"})
        self.assertIn('$kind = $(if ($Demo) { "demo" } else { "addin" })', self.code)

    def test_launcher_passes_demo(self) -> None:
        launcher = LAUNCHER.read_text(encoding="ascii")
        self.assertIn('-File "%SCRIPT%" -Demo %*', launcher)
        self.assertIn('set "SCRIPT=%~dp0Build-Addin.ps1"', launcher)


if __name__ == "__main__":
    unittest.main()
