"""Guard tools/Release.ps1 and the release-check workflow (#222).

PowerShell and Excel do not run in CI, so the script is read as text: the
preflight refuses what #222 lists, a dry run never reaches the tag, the steps
call the existing tools with the release names, and the arguments match what
tools/assemble_release.py and check_release.py expect.
"""
import re
import unittest
from pathlib import Path

import check_release

ROOT = Path(__file__).resolve().parent.parent
SCRIPT = ROOT / "tools" / "Release.ps1"
LAUNCHER = ROOT / "tools" / "Release.cmd"
ASSEMBLER = ROOT / "tools" / "assemble_release.py"
WORKFLOW = ROOT / ".github" / "workflows" / "release-check.yml"


def code_lines(text: str) -> list[str]:
    """Return the script's lines without the comment-based help and comments."""
    text = re.sub(r"<#.*?#>", "", text, flags=re.S)
    return [line for line in text.splitlines() if not line.lstrip().startswith("#")]


class ReleaseScriptTests(unittest.TestCase):
    def setUp(self) -> None:
        self.text = SCRIPT.read_text(encoding="ascii")
        self.code = "\n".join(code_lines(self.text))
        start = self.code.index('$step = Start-Step "1. Preflight"')
        self.preflight = self.code[start:self.code.index('$step = Start-Step "2.', start)]
        self.dry = self.code.index("    if ($DryRun) {")
        self.publish = self.code.index("    else {", self.dry)

    def test_files_are_ascii_for_windows_powershell(self) -> None:
        # Windows PowerShell 5.1 reads a script without a BOM as ANSI.
        SCRIPT.read_bytes().decode("ascii")
        LAUNCHER.read_bytes().decode("ascii")

    def test_preflight_refuses_what_222_lists(self) -> None:
        for needle in (
            'if ($sha -ne $main) {',                                     # not the main commit
            'status --porcelain --untracked-files=normal',               # dirty tree
            "if ($version -notmatch '^\\d+\\.\\d+\\.\\d+$')",            # VERSION
            "\\] - \\d{4}-\\d{2}-\\d{2}\\s*$')) {",                       # dated CHANGELOG section
            '"rev-parse", "-q", "--verify", "refs/tags/$tag"',           # local tag
            '"ls-remote", "--tags", "origin", "refs/tags/$tag"',         # remote tag
            '@("release", "view", $tag',                                 # existing release
            '$_.head_branch -eq "main" -and $_.event -eq "push" -and $_.conclusion -eq "success"',
        ):
            with self.subTest(check=needle):
                self.assertIn(needle, self.preflight)
        self.assertNotIn("--jq", self.code)

    def test_static_gates_come_from_the_ci_report_of_the_commit(self) -> None:
        start = self.code.index('$step = Start-Step "2. Static gates')
        step = self.code[start:self.code.index('$step = Start-Step "3.', start)]
        self.assertIn('@("run", "download", $script:staticRunId,', step)
        self.assertIn('"--name", "static-checks-$sha"', step)
        self.assertIn('$report.status -ne "pass" -or $report.candidate_sha -ne $sha -or $report.dirty -ne $false', step)
        self.assertIn("@($report.checks).Count -ne 11", step)
        self.assertNotIn("check.py", step)
        self.assertNotIn("node", self.code.lower().replace("nodes", ""))

    def test_dry_run_never_tags_or_publishes(self) -> None:
        for needle in ('"tag", "-a", $tag, $sha, "-F", $messagePath', '"push", "origin", "refs/tags/$tag"',
                       '@("release", "create", $tag'):
            with self.subTest(command=needle):
                position = self.code.index(needle)
                self.assertGreater(position, self.publish)
                self.assertEqual(self.code.count(needle), 1)
        self.assertIn("if ($typed -cne $tag) {", self.code[self.publish:])

    def test_origin_main_is_checked_again_before_tagging(self) -> None:
        tagging = self.code.index('"tag", "-a", $tag, $sha')
        recheck = self.code.index('Stop-Release $step "origin/main moved during the run')
        self.assertLess(self.publish, recheck)
        self.assertLess(recheck, tagging)

    def test_steps_use_the_existing_tools_with_release_names(self) -> None:
        for needle in (
            '"Certify-Excel.ps1" @("-BundleDir", $folders.source_standard, "-NoOpen")',
            '"Certify-Excel.ps1" @("-WithUISmoke", "-BundleDir", $folders.source_ui_smoke, "-NoOpen")',
            '"Build-Addin.ps1" @("-PackageName", $addinName, "-BundleDir", $folders.addin_build, "-NoOpen", '
            '"-NoRibbonCheck")',
            '"Build-Addin.ps1" @("-Demo", "-PackageName", $demoName, "-BundleDir", $folders.demo_build, "-NoOpen", '
            '"-NoRibbonCheck")',
            '"Certify-Excel.ps1" @("-Package", $addinPath, "-BundleDir", $folders.addin_run, "-NoOpen")',
            '"Certify-Excel.ps1" @("-Package", $demoPath, "-BundleDir", $folders.demo_run, "-NoOpen")',
            '$addinName = "DATETIMEPICKER.$tag.xlam"',
            '$demoName = "DATETIMEPICKER-demo-$tag.xlsm"',
            '$pdfName = "DATETIMEPICKER-user-manual-$tag.pdf"',
        ):
            with self.subTest(step=needle):
                self.assertIn(needle, self.code)

    def test_child_scripts_share_the_console_and_decide_by_exit_code(self) -> None:
        self.assertIn('Start-Process -FilePath "powershell.exe" -ArgumentList $quoted -NoNewWindow -Wait -PassThru',
                      self.code)
        self.assertIn("return [int]$process.ExitCode", self.code)

    def test_every_required_assembler_argument_is_passed(self) -> None:
        source = ASSEMBLER.read_text(encoding="utf-8")
        required = re.findall(r'add_argument\("(--[a-z-]+)"[^)]*required=True', source)
        self.assertGreater(len(required), 10)
        start = self.code.index("$assemble = @(")
        call = self.code[start:self.code.index("$evidence = Invoke-Python", start)]
        for option in required:
            with self.subTest(option=option):
                self.assertIn(f'"{option}"', call)

    def test_manual_checks_are_the_validators(self) -> None:
        start = self.code.index("$manualChecks = [ordered]@{")
        block = self.code[start:self.code.index("}", start)]
        self.assertEqual(tuple(re.findall(r'^\s+"([a-z-]+)"\s+=', block, flags=re.M)), check_release.MANUAL)
        self.assertIn('status      = "NOT RUN"', self.code)
        self.assertIn("accepted_by = $maintainer", self.code)
        self.assertIn("$maintainer = (& git -C $RepoRoot config user.name | Out-String).Trim()", self.code)

    def test_nothing_is_asked_but_the_tag(self) -> None:
        # A release asks only for the tag; -Title and -EvidenceIssue when missing.
        asked = re.findall(r"Read-Host[^\n]*", self.code)
        self.assertEqual(len(asked), 2, asked)
        self.assertIn('Read-Host "Type $tag to tag $sha and publish the release', self.code)
        self.assertEqual(self.code.count("Read-Required"), 3)
        self.assertNotIn("Read-YesNo", self.code)
        self.assertNotIn("Invoke-Item", self.code)

    def test_published_files_are_downloaded_and_validated_with_the_tag(self) -> None:
        after = self.code[self.code.index('@("release", "create", $tag'):]
        self.assertIn('@("release", "download", $tag', after)
        self.assertIn('"--require-tag"', after)
        self.assertIn('differs from the tested file', after)

    def test_rehearsal_skips_release_state_and_stops_before_evidence(self) -> None:
        guard = self.preflight.index("    if (-not $Rehearse) {")
        for needle in ('if ($sha -ne $main) {', '@("release", "view", $tag', 'head_branch -eq "main"'):
            with self.subTest(check=needle):
                self.assertGreater(self.preflight.index(needle), guard)
        self.assertLess(self.preflight.index("status --porcelain --untracked-files=normal"), guard)
        stop = self.code.index('throw [System.OperationCanceledException]::new("Rehearsal complete after step 7')
        self.assertLess(stop, self.code.index('$step = Start-Step "8. Evidence'))
        self.assertGreater(stop, self.code.index('$step = Start-Step "7. Manual checks"'))
        self.assertIn("$DryRun = $true", self.code[self.code.index("if ($Rehearse) {"):])

    def test_refs_is_a_string(self) -> None:
        # An array parameter does not survive powershell.exe -File.
        self.assertIn('[string]$Refs = ""', self.code)

    def test_launcher_runs_the_script(self) -> None:
        launcher = LAUNCHER.read_text(encoding="ascii")
        self.assertIn('set "SCRIPT=%~dp0Release.ps1"', launcher)
        self.assertIn('-File "%SCRIPT%" %*', launcher)


class ManualAndResumeTests(unittest.TestCase):
    def setUp(self) -> None:
        self.code = "\n".join(code_lines(SCRIPT.read_text(encoding="ascii")))
        start = self.code.index('$step = Start-Step "6. User manual PDF"')
        self.step6 = self.code[start:self.code.index('$step = Start-Step "7.', start)]

    def test_pdf_comes_from_the_ci_run_of_the_commit(self) -> None:
        self.assertIn('"repos/$Repository/actions/workflows/user-manual.yml/runs"', self.step6)
        self.assertIn('($Rehearse -or ($_.event -eq "push" -and $_.head_branch -eq "main"))', self.step6)
        self.assertIn('"--name", "user-manual-$sha"', self.step6)
        self.assertIn('if ((Get-Content -LiteralPath $commitFile -Raw).Trim() -ne $sha) {', self.step6)
        self.assertNotIn("render_manual", self.code)
        self.assertNotIn("import markdown", self.code)

    def test_pdf_is_not_opened_or_asked_about(self) -> None:
        self.assertNotIn("Read-", self.step6)
        self.assertNotIn("Invoke-Item", self.step6)

    def test_a_release_resumes_only_at_the_same_commit(self) -> None:
        self.assertIn('if (-not $Rehearse) {\n            throw "-Resume: that run was at', self.code)
        self.assertIn('if ($previous.tag -ne $tag) {', self.code)
        self.assertIn('if ([bool]$previous.rehearsal -ne [bool]$Rehearse) {', self.code)
        self.assertIn('throw "-Resume: that run already published $tag."', self.code)

    def test_only_a_row_of_passed_steps_is_kept(self) -> None:
        self.assertIn("for ($number = 2; $number -le 8 -and $passed.ContainsKey($number); $number++) {", self.code)
        for number in range(2, 9):
            with self.subTest(step=number):
                self.assertRegex(self.code, rf'if \(-not \(Skip-Step {number} "{number}\. [^"]+"\)\) {{\n'
                                            rf'        \$step = Start-Step "{number}\.')
                self.assertRegex(self.code, rf"(?m)^    {number} = @\(")
        self.assertNotIn("Skip-Step 1 ", self.code)
        self.assertNotIn("Skip-Step 9 ", self.code)

    def test_files_to_publish_are_checked_again_after_a_kept_step_8(self) -> None:
        # Step 9 never runs inside Skip-Step 8, so a resumed run re-checks the
        # assets folder before the dry run ends and before the confirmation.
        start = self.code.index('$step = Start-Step "9. Final confirmation"')
        recheck = self.code[start:self.code.index("    if ($DryRun) {", start)]
        self.assertGreater(start, self.code.index('if (-not (Skip-Step 8 "8.'))
        self.assertIn('"--evidence", $recordPath, "--assets-dir", $assetsDir', recheck)
        self.assertNotIn("--require-tag", recheck)
        self.assertIn('if ($recheck.Code -ne 0) { Stop-Release $step', recheck)
        self.assertIn('(Get-Sha256 (Join-Path $folders.evidence "release-evidence.json"))', recheck)
        self.assertIn("(Get-Sha256 $pdfPath) -ne [string]$documents[0].sha256", recheck)

    def test_a_step_that_runs_again_moves_its_outputs_aside(self) -> None:
        self.assertIn('$aside = $path + ".failed-" + $stamp', self.code)
        self.assertIn("Move-Item -LiteralPath $path -Destination $aside", self.code)


class UserManualWorkflowTests(unittest.TestCase):
    def setUp(self) -> None:
        self.text = (ROOT / ".github" / "workflows" / "user-manual.yml").read_text(encoding="utf-8")

    def test_runs_where_release_and_rehearsal_look(self) -> None:
        self.assertRegex(self.text, r'(?m)^  push:\n    branches:\n      - main\n      - "release/\*\*"$')
        self.assertRegex(self.text, r"(?m)^  pull_request:$")

    def test_renders_and_names_the_exact_commit(self) -> None:
        self.assertIn("COMMIT: ${{ github.event.pull_request.head.sha || github.sha }}", self.text)
        self.assertIn("ref: ${{ env.COMMIT }}", self.text)
        self.assertIn('[[ "$(git rev-parse HEAD)" == "$COMMIT" ]]', self.text)
        self.assertIn('python tools/render_manual.py --ref "$COMMIT"', self.text)
        self.assertIn("name: user-manual-${{ env.COMMIT }}", self.text)
        self.assertIn("--require-hashes --only-binary=:all: -r tools/requirements-docs.txt", self.text)


class ReleaseCheckWorkflowTests(unittest.TestCase):
    def setUp(self) -> None:
        self.text = WORKFLOW.read_text(encoding="utf-8")

    def test_runs_when_a_release_is_published_and_by_hand(self) -> None:
        self.assertRegex(self.text, r"(?m)^  release:\n    types: \[published\]$")
        self.assertRegex(self.text, r"(?m)^  workflow_dispatch:$")

    def test_release_event_reads_the_tag_and_the_published_record(self) -> None:
        self.assertIn('RELEASE_TAG_INPUT: ${{ github.event.release.tag_name }}', self.text)
        self.assertIn('candidate="$(git rev-parse "refs/tags/$RELEASE_TAG_INPUT^{commit}")"', self.text)
        self.assertIn("--pattern release-evidence.json", self.text)
        self.assertIn("--require-tag", self.text)

    def test_untrusted_values_reach_the_shell_only_through_the_environment(self) -> None:
        for line in self.text.splitlines():
            if "${{" in line and ("github.event." in line or "inputs." in line):
                with self.subTest(line=line.strip()):
                    self.assertRegex(line, r"^\s+[A-Z_]+: \$\{\{")


if __name__ == "__main__":
    unittest.main()
