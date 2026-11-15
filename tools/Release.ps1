<#
.SYNOPSIS
Run the whole release procedure of RELEASING.md from one command on the
Windows host: certify the exact main commit, build and test both packages,
assemble the evidence, then tag and publish (#222).

.DESCRIPTION
On a clean checkout of the exact origin/main commit, in order, stopping at the
first failure:
 1. preflight: origin/main is fetched and equals the checkout; the tree is
    clean; VERSION is X.Y.Z with a dated CHANGELOG section; the tag and the
    GitHub Release do not exist yet; the static-checks workflow has a
    successful push run on main for this commit; GitHub CLI is signed in;
    Python 3.11+ with the documentation packages is on the PATH;
 2. the static gates: the static-checks.json of that push run, downloaded
    from GitHub Actions (the report release-check.yml also reads), so the
    eleven gates are those CI ran on this exact commit; and the Ribbon check
    (.github/scripts/check-ribbon-resources.py), run here;
 3. the source runs: tools/Certify-Excel.ps1, standard and -WithUISmoke;
 4. the add-in and demo builds: tools/Build-Addin.ps1 and -Demo, with the
    release file names, with -NoRibbonCheck: nothing is asked;
 5. both package runs: tools/Certify-Excel.ps1 -Package on each package;
 6. the user manual PDF: the artifact the User manual PDF workflow
    (.github/workflows/user-manual.yml) rendered for this commit on its push
    run on main, downloaded with GitHub CLI and checked against the commit. The
    release host's managed Chrome and Edge refuse the DevTools connection that
    headless printing needs, so the PDF is printed by CI;
 7. the six manual checks of RELEASING section 8, recorded as NOT RUN and
    accepted by the Git user.name of the checkout: nothing is asked, and the
    release notes list them as accepted gaps;
 8. the evidence (tools/assemble_release.py): release-evidence.json, the tag
    message and the release notes, validated with tools/check_release.py;
 9. a final confirmation: you type the tag to publish;
10. the annotated tag, pushed, and the GitHub Release with its assets
    (GitHub CLI). The published files are downloaded again and validated
    with check_release.py --require-tag; the Verify published release
    evidence workflow then runs on its own.

-DryRun does steps 1 to 9 and stops before the tag: nothing is tagged,
pushed or published.

-Rehearse tries the command on any clean checkout, such as a pull request
branch, before a release is prepared. It skips the release-state checks of
the preflight (main, VERSION, CHANGELOG, tag, release, static run), takes the
PDF from any successful run for the commit, runs steps 2 to 7 and stops
before the evidence: its result is never release evidence.

-Resume <release folder> continues a run that stopped, in the same folder.
The commit, tag and mode must be the same. The preflight runs again; the
steps from 2 that passed before, in a row, are kept, and the run continues
from the first one that did not. A step that runs again moves its earlier
folders aside (.failed-<time>), because the tools never overwrite evidence.
tools/assemble_release.py then cross-checks everything that was kept, and
step 9 always checks the packages, the PDF and release-evidence.json in the
assets folder against the evidence again, so files changed after a kept
step 8 are never published. Every step writes into one release folder next
to the checkout, DATETIMEPICKER-release-vX.Y.Z-<sha12>-<yyyyMMdd-HHmmss>, which ends
with release.txt and release.json.

.REQUIREMENTS
- Windows with desktop Excel, Git, Python 3.11 or later and GitHub CLI (gh,
  signed in with access to the repository) on the PATH. Node.js is not
  needed: the static report comes from CI.
- "Trust access to the VBA project object model" already on. This script
  never changes it.
- Close other Excel work first, and do not use Excel while a step runs.

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Release.ps1 -DryRun -Title "User-facing consistency and Ribbon source patch" -Refs 88,89 -EvidenceIssue 122
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$Title,
    [Parameter(Mandatory = $false)]
    [string]$Refs = "",
    [Parameter(Mandatory = $false)]
    [int]$EvidenceIssue = 0,
    [Parameter(Mandatory = $false)]
    [string]$Repository = "danielep71/VBA-DATETIMEPICKER",
    [Parameter(Mandatory = $false)]
    [string]$RepoRoot,
    [string]$Resume,
    [switch]$DryRun,
    [switch]$Rehearse
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $RepoRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
}
$RepoRoot = (Resolve-Path -LiteralPath $RepoRoot).Path
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$script:sessionLog = $null
$manualChecks = [ordered]@{
    "provider-ownership"      = "second-provider refusal, owner teardown, clean-session DP_Show/DP_Close"
    "clock-lifecycle"         = "real OnTime delivery, rename/requalification, close and restart"
    "grid-ownership"          = "grid creation/removal, caller shapes preserved"
    "application-state"       = "events/settings restored; write-back, tables, formulas, protection/array refusal"
    "packaged-entry-paths"    = "keyboard, context menu, grid, Ribbon, startup; visuals/images and callbacks in both packages"
    "ribbon-image-provenance" = "tracked image sources and redistribution evidence, distinct from visual rendering"
}

function Get-Timestamp {
    return (Get-Date).ToString("yyyy-MM-ddTHH:mm:sszzz")
}

function Write-Session {
    param([string]$Text)
    $line = "{0}  {1}" -f (Get-Date).ToString("HH:mm:ss"), $Text
    Write-Host $line
    if ($null -ne $script:sessionLog) {
        [System.IO.File]::AppendAllText($script:sessionLog, $line + "`r`n", $script:utf8NoBom)
    }
}

function Invoke-Native {
    # Runs a program and returns its exit code and output lines, without
    # stopping on stderr. Python writes and is read as UTF-8.
    param([string]$Program, [string[]]$Arguments)
    $previous = $ErrorActionPreference
    $previousUtf8 = $env:PYTHONUTF8
    $ErrorActionPreference = "Continue"
    $env:PYTHONUTF8 = "1"
    try {
        $output = @(& $Program @Arguments 2>&1 | ForEach-Object { [string]$_ })
        $code = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $previous
        $env:PYTHONUTF8 = $previousUtf8
    }
    return [pscustomobject]@{ Code = $code; Output = $output }
}

function Find-Python {
    # The python command, or the py launcher, at version 3.11 or later.
    foreach ($candidate in @(
            [pscustomobject]@{ Program = "python"; Prefix = @() },
            [pscustomobject]@{ Program = "py"; Prefix = @("-3") })) {
        if ($null -eq (Get-Command $candidate.Program -ErrorAction SilentlyContinue)) {
            continue
        }
        $probe = Invoke-Native -Program $candidate.Program -Arguments ($candidate.Prefix + @("--version"))
        $text = ($probe.Output -join " ")
        if ($probe.Code -eq 0 -and $text -match 'Python 3\.(\d+)\.' -and [int]$Matches[1] -ge 11) {
            return [pscustomobject]@{ Program = $candidate.Program; Prefix = $candidate.Prefix; Version = $text.Trim() }
        }
    }
    return $null
}

function Invoke-Python {
    param([object]$Python, [string[]]$Arguments)
    return Invoke-Native -Program $Python.Program -Arguments ($Python.Prefix + $Arguments)
}

function Read-Required {
    param([string]$Prompt)
    while ($true) {
        $answer = (Read-Host $Prompt).Trim()
        if ($answer.Length -gt 0) { return $answer }
    }
}

function Get-Sha256 {
    param([string]$Path)
    return (Get-FileHash -Algorithm SHA256 -LiteralPath $Path).Hash.ToLowerInvariant()
}

function Start-Step {
    param([string]$Name)
    Write-Host ""
    Write-Session "=== $Name"
    $step = [ordered]@{ name = $Name; status = "STARTED"; started_at = Get-Timestamp; finished_at = $null; detail = $null }
    $script:summary.steps += $step
    return $step
}

function Complete-Step {
    param([System.Collections.IDictionary]$Step, [string]$Detail)
    $Step.status = "PASS"
    $Step.detail = $Detail
    $Step.finished_at = Get-Timestamp
    Write-Session "PASS: $($Step.name)$(if ($Detail) { ' - ' + $Detail } else { '' })"
}

function Stop-Release {
    param([System.Collections.IDictionary]$Step, [string]$Reason)
    $Step.status = "FAIL"
    $Step.detail = $Reason
    $Step.finished_at = Get-Timestamp
    throw "$($Step.name): $Reason"
}

function Invoke-ToolScript {
    # Runs a tools\*.ps1 script in its own PowerShell process that shares this
    # console, so its questions reach the operator and its output is not
    # captured; its own exit code decides. Values with spaces are quoted;
    # switch names are passed bare, so that they bind as switches.
    param([string]$Script, [string[]]$Arguments)
    $path = Join-Path $RepoRoot "tools\$Script"
    Write-Session "Running $Script $($Arguments -join ' ')"
    $quoted = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ('"' + $path + '"')) +
        @($Arguments | ForEach-Object { if ($_ -match '\s') { '"' + $_ + '"' } else { $_ } })
    $process = Start-Process -FilePath "powershell.exe" -ArgumentList $quoted -NoNewWindow -Wait -PassThru
    Write-Session "$Script exit code $($process.ExitCode)"
    return [int]$process.ExitCode
}

# ----------------------------------------------------------------------------
# Inputs and release folder
# ----------------------------------------------------------------------------
$sha = (& git -C $RepoRoot rev-parse HEAD).Trim()
if ($LASTEXITCODE -ne 0 -or $sha -notmatch '^[0-9a-f]{40}$') {
    throw "Cannot read the checked-out commit of $RepoRoot. Is Git on the PATH?"
}
$version = (Get-Content -LiteralPath (Join-Path $RepoRoot "VERSION") -Raw).Trim()
$tag = "v$version"
$stamp = (Get-Date).ToString("yyyyMMdd-HHmmss")
$previous = $null
$script:resumedFrom = $null
if (-not [string]::IsNullOrWhiteSpace($Resume)) {
    # Continue a run that stopped: same folder, same commit, tag and mode.
    $releaseDir = (Resolve-Path -LiteralPath $Resume).Path
    $previousFile = Join-Path $releaseDir "release.json"
    if (-not (Test-Path -LiteralPath $previousFile)) {
        throw "-Resume needs a release folder with release.json: $releaseDir"
    }
    $previous = Get-Content -LiteralPath $previousFile -Raw -Encoding UTF8 | ConvertFrom-Json
    if ([bool]$previous.rehearsal -ne [bool]$Rehearse) {
        throw "-Resume: that run $(if ($previous.rehearsal) { 'was' } else { 'was not' }) a rehearsal; resume it the same way."
    }
    if ($previous.tag -ne $tag) {
        throw "-Resume: that run was for $($previous.tag), not $tag."
    }
    if ($previous.candidate_sha -ne $sha) {
        # Release evidence is bound to one commit. A rehearsal is never
        # evidence, so it may keep steps run at an earlier commit of the tools.
        if (-not $Rehearse) {
            throw "-Resume: that run was at $($previous.candidate_sha), not $sha; a release resumes only at the same commit."
        }
        $script:resumedFrom = [string]$previous.candidate_sha
    }
    if ($previous.published) {
        throw "-Resume: that run already published $tag."
    }
    if ([string]::IsNullOrWhiteSpace($Title)) { $Title = [string]$previous.title }
    if ($EvidenceIssue -le 0) { $EvidenceIssue = [int]$previous.evidence }
    if ([string]::IsNullOrWhiteSpace($Refs)) { $Refs = (@($previous.refs) -join ",") }
}
else {
    $releaseDir = Join-Path (Split-Path -Parent $RepoRoot) ("DATETIMEPICKER-release-$tag-" + $sha.Substring(0, 12) + "-" + $stamp)
    New-Item -ItemType Directory -Path $releaseDir | Out-Null
    $releaseDir = (Resolve-Path -LiteralPath $releaseDir).Path
}
if ([string]::IsNullOrWhiteSpace($Title)) {
    $Title = Read-Required "Release title, one line (for example: User-facing consistency and Ribbon source patch)"
}
if ($Title -match "[`r`n]") {
    throw "-Title must be one line."
}
# -Refs is one string, such as "88,89": an array does not survive
# powershell.exe -File.
$refNumbers = @()
foreach ($item in ($Refs -split '[,\s]+' | Where-Object { $_ })) {
    if ($item -notmatch '^#?\d+$') {
        throw "-Refs takes issue numbers separated by commas, such as 88,89: '$item'"
    }
    $refNumbers += [int]($item.TrimStart('#'))
}
if ($EvidenceIssue -le 0) {
    $EvidenceIssue = [int](Read-Required "Number of the certification issue (for example 122)")
}
$assetsDir = Join-Path $releaseDir "release-assets"
New-Item -ItemType Directory -Path $assetsDir -Force | Out-Null
$script:sessionLog = Join-Path $releaseDir "session.log"
$addinName = "DATETIMEPICKER.$tag.xlam"
$demoName = "DATETIMEPICKER-demo-$tag.xlsm"
$pdfName = "DATETIMEPICKER-user-manual-$tag.pdf"
$manualPath = Join-Path $releaseDir "manual-checks.json"
$pdfPath = Join-Path $assetsDir $pdfName
$folders = [ordered]@{
    static          = Join-Path $releaseDir "static"
    source_standard = Join-Path $releaseDir "source-standard"
    source_ui_smoke = Join-Path $releaseDir "source-ui_smoke"
    addin_build     = Join-Path $releaseDir "addin-build"
    demo_build      = Join-Path $releaseDir "demo-build"
    addin_run       = Join-Path $releaseDir "addin-run"
    demo_run        = Join-Path $releaseDir "demo-run"
    manual          = Join-Path $releaseDir "manual"
    evidence        = Join-Path $releaseDir "evidence"
    published       = Join-Path $releaseDir "published"
}
$addinPath = Join-Path $folders.addin_build $addinName
$demoPath = Join-Path $folders.demo_build $demoName
# What each step writes; a step that runs again on -Resume moves these aside.
$stepOutputs = @{
    2 = @($folders.static)
    3 = @($folders.source_standard, $folders.source_ui_smoke)
    4 = @($folders.addin_build, $folders.demo_build)
    5 = @($folders.addin_run, $folders.demo_run)
    6 = @($folders.manual, $pdfPath)
    7 = @($manualPath)
    8 = @($folders.evidence)
}

$script:summary = [ordered]@{
    tool          = "tools/Release.ps1"
    candidate_sha = $sha
    version       = $version
    tag           = $tag
    title         = $Title
    refs          = $refNumbers
    evidence      = $EvidenceIssue
    dry_run       = [bool]$DryRun
    rehearsal     = [bool]$Rehearse
    started_at    = Get-Timestamp
    finished_at   = $null
    static_run    = $null
    steps         = @()
    published     = $false
    result        = "FAIL"
    reason        = $null
    resumed       = ($null -ne $previous)
    resumed_from  = $script:resumedFrom
    kept_steps    = @()
}
if ($Rehearse) {
    $DryRun = $true
    $script:summary.dry_run = $true
}
Write-Session ("Release $tag of $sha" + $(if ($Rehearse) { " (REHEARSAL: steps 2 to 7, not release evidence)" }
        elseif ($DryRun) { " (dry run: no tag, no release)" } else { "" }))
Write-Session "Release folder: $releaseDir"

# The steps from 2 that passed in the earlier run, in a row, are kept.
$kept = @()
if ($null -ne $previous) {
    $passed = @{}
    foreach ($item in @($previous.steps)) {
        if ([string]$item.status -eq "PASS" -and [string]$item.name -match '^(\d+)\.') {
            $passed[[int]$Matches[1]] = [string]$item.detail
        }
    }
    for ($number = 2; $number -le 8 -and $passed.ContainsKey($number); $number++) {
        $kept += $number
    }
    $script:previousDetails = $passed
    Write-Session "Resuming: steps kept from the earlier run: $(if ($kept.Count) { $kept -join ', ' } else { 'none' })"
    if ($script:resumedFrom) {
        Write-Session "Rehearsal: the kept steps ran at $($script:resumedFrom), not at $sha"
    }
}

function Skip-Step {
    # True when the step is kept from the earlier run. Otherwise, on -Resume,
    # its earlier outputs are moved aside so that the tools can write afresh.
    param([int]$Number, [string]$Name)
    if ($kept -contains $Number) {
        $script:summary.steps += [ordered]@{ name = $Name; status = "PASS"; started_at = $null; finished_at = $null
            detail = "kept from the earlier run: $($script:previousDetails[$Number])" }
        $script:summary.kept_steps += $Number
        Write-Session "Kept: $Name"
        return $true
    }
    if ($null -ne $previous) {
        foreach ($path in $stepOutputs[$Number]) {
            if (Test-Path -LiteralPath $path) {
                $aside = $path + ".failed-" + $stamp
                Move-Item -LiteralPath $path -Destination $aside
                Write-Session "Moved the earlier $([System.IO.Path]::GetFileName($path)) aside to $([System.IO.Path]::GetFileName($aside))"
            }
        }
    }
    return $false
}

$python = $null
try {
    # ------------------------------------------------------------------------
    # 1. Preflight
    # ------------------------------------------------------------------------
    $step = Start-Step "1. Preflight"
    $changes = (& git -C $RepoRoot status --porcelain --untracked-files=normal | Out-String).Trim()
    if ($changes.Length -gt 0) { Stop-Release $step "The working tree is not clean: $changes" }
    if ($null -eq (Get-Command gh -ErrorAction SilentlyContinue)) { Stop-Release $step "GitHub CLI (gh) is not on the PATH." }
    $auth = Invoke-Native -Program "gh" -Arguments @("auth", "status")
    if ($auth.Code -ne 0) { Stop-Release $step "GitHub CLI is not signed in: run gh auth login." }
    if (-not $Rehearse) {
        $fetch = Invoke-Native -Program "git" -Arguments @("-C", $RepoRoot, "fetch", "origin", "main", "--tags")
        if ($fetch.Code -ne 0) { Stop-Release $step "git fetch origin main failed: $($fetch.Output -join ' ')" }
        $main = (& git -C $RepoRoot rev-parse "refs/remotes/origin/main").Trim()
        if ($sha -ne $main) {
            Stop-Release $step "The checkout is $sha, but origin/main is $main. Check out the exact main commit."
        }
        if ($version -notmatch '^\d+\.\d+\.\d+$') { Stop-Release $step "VERSION is '$version', not X.Y.Z." }
        $changelog = Get-Content -LiteralPath (Join-Path $RepoRoot "CHANGELOG.md") -Raw -Encoding UTF8
        if ($changelog -notmatch ('(?m)^## \[' + [regex]::Escape($version) + '\] - \d{4}-\d{2}-\d{2}\s*$')) {
            Stop-Release $step "CHANGELOG.md has no dated [$version] section; VERSION and CHANGELOG disagree."
        }
        $localTag = Invoke-Native -Program "git" -Arguments @("-C", $RepoRoot, "rev-parse", "-q", "--verify", "refs/tags/$tag")
        if ($localTag.Code -eq 0) { Stop-Release $step "The tag $tag already exists." }
        $remoteTag = Invoke-Native -Program "git" -Arguments @("-C", $RepoRoot, "ls-remote", "--tags", "origin", "refs/tags/$tag")
        if ($remoteTag.Code -ne 0 -or ($remoteTag.Output -join "").Trim().Length -gt 0) {
            Stop-Release $step "The tag $tag exists on origin, or origin could not be read."
        }
        $existing = Invoke-Native -Program "gh" -Arguments @("release", "view", $tag, "--repo", $Repository)
        if ($existing.Code -eq 0) { Stop-Release $step "A GitHub Release $tag already exists." }
        # The JSON is filtered here: a --jq filter would need double quotes, which
        # Windows PowerShell 5.1 does not pass intact to a native program.
        $runs = Invoke-Native -Program "gh" -Arguments @("api", "-X", "GET",
            "repos/$Repository/actions/workflows/static-checks.yml/runs",
            "-f", "head_sha=$sha", "-f", "event=push", "-f", "status=success", "-f", "per_page=100")
        $runUrl = ""
        $script:staticRunId = $null
        if ($runs.Code -eq 0) {
            $found = @((($runs.Output -join "`n") | ConvertFrom-Json).workflow_runs | Where-Object {
                    $_.head_sha -eq $sha -and $_.head_branch -eq "main" -and $_.event -eq "push" -and $_.conclusion -eq "success" })
            if ($found.Count -gt 0) {
                $runUrl = [string]$found[0].html_url
                $script:staticRunId = [string]$found[0].id
            }
        }
        if ($runUrl.Length -eq 0) {
            Stop-Release $step "No successful static-checks push run on main for $sha (RELEASING section 1, step 5)."
        }
        $script:summary.static_run = $runUrl
    }
    $python = Find-Python
    if ($null -eq $python) { Stop-Release $step "Python 3.11 or later was not found (python or py -3)." }
    Complete-Step $step $(if ($Rehearse) { "rehearsal: clean tree, gh signed in, $($python.Version); release state not checked" }
        else { "main $sha, clean, $tag new, static run $($script:summary.static_run), $($python.Version)" })

    # ------------------------------------------------------------------------
    # 2. Static gates and the Ribbon check
    # ------------------------------------------------------------------------
    if (-not (Skip-Step 2 "2. Static gates and Ribbon check")) {
        $step = Start-Step "2. Static gates and Ribbon check"
        New-Item -ItemType Directory -Path $folders.static | Out-Null
        $staticDetail = "rehearsal: static report not retrieved"
        if (-not $Rehearse) {
            # The report CI wrote on the push run of this exact commit, as
            # release-check.yml reads it after publication.
            $download = Invoke-Native -Program "gh" -Arguments @("run", "download", $script:staticRunId,
                "--repo", $Repository, "--name", "static-checks-$sha", "--dir", $folders.static)
            $staticReport = Join-Path $folders.static "static-checks.json"
            if ($download.Code -ne 0 -or -not (Test-Path -LiteralPath $staticReport)) {
                Stop-Release $step "The static-checks report of run $($script:staticRunId) could not be downloaded (artifacts expire after 30 days)."
            }
            $report = Get-Content -LiteralPath $staticReport -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($report.status -ne "pass" -or $report.candidate_sha -ne $sha -or $report.dirty -ne $false -or
                $report.mode -ne "committed" -or @($report.checks).Count -ne 11) {
                Stop-Release $step "The static report of run $($script:staticRunId) is not a clean committed PASS of 11 gates for $sha."
            }
            $staticDetail = "11 gates PASS in CI run $($script:staticRunId)"
        }
        $ribbon = Invoke-Python -Python $python -Arguments @((Join-Path $RepoRoot ".github\scripts\check-ribbon-resources.py"),
            "--root", (Join-Path $RepoRoot "src\ribbon"))
        [System.IO.File]::WriteAllLines((Join-Path $folders.static "ribbon-resources.txt"),
            [string[]](@("Commit: $sha", "Command: python .github/scripts/check-ribbon-resources.py --root src/ribbon",
                         "Exit code: $($ribbon.Code)", "") + $ribbon.Output), $utf8NoBom)
        if ($ribbon.Code -ne 0) { Stop-Release $step "check-ribbon-resources.py failed." }
        Complete-Step $step "$staticDetail; Ribbon resources pass"
    }

    # ------------------------------------------------------------------------
    # 3. Source runs
    # ------------------------------------------------------------------------
    if (-not (Skip-Step 3 "3. Source runs")) {
        $step = Start-Step "3. Source runs"
        if ((Invoke-ToolScript "Certify-Excel.ps1" @("-BundleDir", $folders.source_standard, "-NoOpen")) -ne 0) {
            Stop-Release $step "The standard source run failed; see $($folders.source_standard)."
        }
        if ((Invoke-ToolScript "Certify-Excel.ps1" @("-WithUISmoke", "-BundleDir", $folders.source_ui_smoke, "-NoOpen")) -ne 0) {
            Stop-Release $step "The UI-smoke source run failed; see $($folders.source_ui_smoke)."
        }
        Complete-Step $step "standard and UI smoke PASS"
    }

    # ------------------------------------------------------------------------
    # 4. Package builds
    # ------------------------------------------------------------------------
    if (-not (Skip-Step 4 "4. Add-in and demo builds")) {
        $step = Start-Step "4. Add-in and demo builds"
        if ((Invoke-ToolScript "Build-Addin.ps1" @("-PackageName", $addinName, "-BundleDir", $folders.addin_build, "-NoOpen", "-NoRibbonCheck")) -ne 0) {
            Stop-Release $step "The add-in build failed; see $($folders.addin_build)."
        }
        if ((Invoke-ToolScript "Build-Addin.ps1" @("-Demo", "-PackageName", $demoName, "-BundleDir", $folders.demo_build, "-NoOpen", "-NoRibbonCheck")) -ne 0) {
            Stop-Release $step "The demo build failed; see $($folders.demo_build)."
        }
        Complete-Step $step "$addinName $(Get-Sha256 $addinPath), $demoName $(Get-Sha256 $demoPath)"
    }

    # ------------------------------------------------------------------------
    # 5. Package runs
    # ------------------------------------------------------------------------
    if (-not (Skip-Step 5 "5. Package runs")) {
        $step = Start-Step "5. Package runs"
        if ((Invoke-ToolScript "Certify-Excel.ps1" @("-Package", $addinPath, "-BundleDir", $folders.addin_run, "-NoOpen")) -ne 0) {
            Stop-Release $step "The add-in package run failed; see $($folders.addin_run)."
        }
        if ((Invoke-ToolScript "Certify-Excel.ps1" @("-Package", $demoPath, "-BundleDir", $folders.demo_run, "-NoOpen")) -ne 0) {
            Stop-Release $step "The demo package run failed; see $($folders.demo_run)."
        }
        Complete-Step $step "both packs PASS in both packages"
    }

    # ------------------------------------------------------------------------
    # 6. User manual PDF
    # ------------------------------------------------------------------------
    if (-not (Skip-Step 6 "6. User manual PDF")) {
        $step = Start-Step "6. User manual PDF"
        # Printed by CI: the release host's managed Chrome and Edge refuse the
        # DevTools connection that headless printing needs. A release takes the
        # push run on main for this commit; a rehearsal takes any run for it.
        $manualRuns = Invoke-Native -Program "gh" -Arguments @("api", "-X", "GET",
            "repos/$Repository/actions/workflows/user-manual.yml/runs",
            "-f", "head_sha=$sha", "-f", "status=success", "-f", "per_page=100")
        $manualRun = $null
        if ($manualRuns.Code -eq 0) {
            $manualRun = @((($manualRuns.Output -join "`n") | ConvertFrom-Json).workflow_runs | Where-Object {
                    $_.head_sha -eq $sha -and $_.conclusion -eq "success" -and
                    ($Rehearse -or ($_.event -eq "push" -and $_.head_branch -eq "main")) }) | Select-Object -First 1
        }
        if ($null -eq $manualRun) {
            Stop-Release $step ("No successful User manual PDF run for $sha" +
                $(if ($Rehearse) { "; push the branch or open its pull request first." } else { " on a push to main." }))
        }
        New-Item -ItemType Directory -Path $folders.manual | Out-Null
        $download = Invoke-Native -Program "gh" -Arguments @("run", "download", [string]$manualRun.id,
            "--repo", $Repository, "--name", "user-manual-$sha", "--dir", $folders.manual)
        $printed = Join-Path $folders.manual "USER_MANUAL.pdf"
        $commitFile = Join-Path $folders.manual "COMMIT"
        if ($download.Code -ne 0 -or -not (Test-Path -LiteralPath $printed) -or -not (Test-Path -LiteralPath $commitFile)) {
            Stop-Release $step "The PDF of run $($manualRun.id) could not be downloaded (artifacts expire after 90 days)."
        }
        if ((Get-Content -LiteralPath $commitFile -Raw).Trim() -ne $sha) {
            Stop-Release $step "The PDF of run $($manualRun.id) was not rendered from $sha."
        }
        Copy-Item -LiteralPath $printed -Destination $pdfPath
        Write-Session "PDF from $($manualRun.html_url)"
        Complete-Step $step "$pdfName $(Get-Sha256 $pdfPath) from CI run $($manualRun.id)"
    }

    # ------------------------------------------------------------------------
    # 7. Manual checks
    # ------------------------------------------------------------------------
    if (-not (Skip-Step 7 "7. Manual checks")) {
        $step = Start-Step "7. Manual checks"
        # Nothing is asked: the checks are recorded as not run, accepted by the
        # maintainer running the release, and the notes list them as gaps.
        $maintainer = (& git -C $RepoRoot config user.name | Out-String).Trim()
        if ([string]::IsNullOrWhiteSpace($maintainer)) { Stop-Release $step "Git user.name is not set; it names who accepts the manual checks not run." }
        $manual = [ordered]@{}
        foreach ($name in $manualChecks.Keys) {
            $manual[$name] = [ordered]@{
                status      = "NOT RUN"
                detail      = "Not run for this release ($($manualChecks[$name])); the automated packs ran in both packages"
                accepted_by = $maintainer
            }
        }
        [System.IO.File]::WriteAllText($manualPath, ($manual | ConvertTo-Json -Depth 4), $utf8NoBom)
        Complete-Step $step ("$(@($manual.Values | Where-Object { $_.status -eq 'PASS' }).Count) PASS, " +
                             "$(@($manual.Values | Where-Object { $_.status -eq 'NOT RUN' }).Count) NOT RUN")
    }

    if ($Rehearse) {
        # A rehearsal runs on a commit that is not a release candidate, so
        # check_release.py would rightly refuse its evidence.
        $script:summary.result = "PASS"
        throw [System.OperationCanceledException]::new("Rehearsal complete after step 7; no evidence, tag or release.")
    }

    # ------------------------------------------------------------------------
    # 8. Evidence
    # ------------------------------------------------------------------------
    if (-not (Skip-Step 8 "8. Evidence, tag message and release notes")) {
        $step = Start-Step "8. Evidence, tag message and release notes"
        foreach ($pair in @(@($addinPath, $addinName), @($demoPath, $demoName))) {
            $target = Join-Path $assetsDir $pair[1]
            Copy-Item -LiteralPath $pair[0] -Destination $target
            if ((Get-Sha256 $target) -ne (Get-Sha256 $pair[0])) { Stop-Release $step "$($pair[1]) changed while it was copied." }
        }
        $assemble = @((Join-Path $RepoRoot "tools\assemble_release.py"), "--root", $RepoRoot, "--candidate-sha", $sha,
            "--assets-dir", $assetsDir, "--static-report", (Join-Path $folders.static "static-checks.json"),
            "--source-standard", $folders.source_standard, "--source-ui-smoke", $folders.source_ui_smoke,
            "--addin-build", $folders.addin_build, "--addin-run", $folders.addin_run,
            "--demo-build", $folders.demo_build, "--demo-run", $folders.demo_run,
            "--manual", $manualPath, "--title", $Title, "--evidence-issue", [string]$EvidenceIssue,
            "--output-dir", $folders.evidence)
        if ($refNumbers.Count -gt 0) { $assemble += @("--refs") + @($refNumbers | ForEach-Object { [string]$_ }) }
        $evidence = Invoke-Python -Python $python -Arguments $assemble
        foreach ($line in $evidence.Output) { Write-Session "  $line" }
        if ($evidence.Code -ne 0) { Stop-Release $step "tools/assemble_release.py refused the evidence." }
        Copy-Item -LiteralPath (Join-Path $folders.evidence "release-evidence.json") -Destination $assetsDir
        Complete-Step $step "release-evidence.json validated: $(($evidence.Output | Select-Object -Last 1))"
    }

    # ------------------------------------------------------------------------
    # 9. Final confirmation
    # ------------------------------------------------------------------------
    $step = Start-Step "9. Final confirmation"
    $messagePath = Join-Path $folders.evidence "tag-message.txt"
    $notesPath = Join-Path $folders.evidence "release-notes.md"
    # Step 8 may have been kept from an earlier run, so the files that will be
    # published are checked again against the evidence, every time: a package,
    # the PDF or the record changed since then is never tagged or published.
    $recordPath = Join-Path $assetsDir "release-evidence.json"
    if ((Get-Sha256 $recordPath) -ne (Get-Sha256 (Join-Path $folders.evidence "release-evidence.json"))) {
        Stop-Release $step "release-evidence.json in $assetsDir differs from the assembled record."
    }
    $recheck = Invoke-Python -Python $python -Arguments @((Join-Path $RepoRoot "tools\check_release.py"),
        "--root", $RepoRoot, "--evidence", $recordPath, "--assets-dir", $assetsDir,
        "--static-report", (Join-Path $folders.static "static-checks.json"),
        "--output", (Join-Path $releaseDir "assets-check-$stamp.json"))
    foreach ($line in $recheck.Output) { Write-Session "  $line" }
    if ($recheck.Code -ne 0) { Stop-Release $step "check_release.py refused the files in $assetsDir." }
    $documents = @((Get-Content -LiteralPath $recordPath -Raw -Encoding UTF8 | ConvertFrom-Json).documents)
    if ($documents.Count -ne 1 -or $documents[0].filename -ne $pdfName -or
        (Get-Sha256 $pdfPath) -ne [string]$documents[0].sha256) {
        Stop-Release $step "$pdfName differs from the PDF recorded in release-evidence.json."
    }
    Write-Session "The packages, the PDF and release-evidence.json in $assetsDir match the evidence."
    Write-Host ""
    Write-Host "Tag message:"
    Write-Host ""
    Get-Content -LiteralPath $messagePath -Encoding UTF8 | ForEach-Object { Write-Host "  $_" }
    Write-Host ""
    Write-Host "Release notes: $notesPath"
    Write-Host "Assets:        $assetsDir"
    if ($DryRun) {
        Complete-Step $step "dry run: nothing tagged or published"
        $script:summary.result = "PASS"
    }
    else {
        $typed = (Read-Host "Type $tag to tag $sha and publish the release, anything else to stop").Trim()
        if ($typed -cne $tag) { Stop-Release $step "The operator did not confirm; nothing was tagged or published." }
        Complete-Step $step "confirmed by the operator"

        # --------------------------------------------------------------------
        # 10. Tag, release and upload
        # --------------------------------------------------------------------
        $step = Start-Step "10. Tag, GitHub Release and verification"
        $again = Invoke-Native -Program "git" -Arguments @("-C", $RepoRoot, "fetch", "origin", "main", "--tags")
        if ($again.Code -ne 0 -or (& git -C $RepoRoot rev-parse "refs/remotes/origin/main").Trim() -ne $sha) {
            Stop-Release $step "origin/main moved during the run; nothing was tagged or published."
        }
        $tagged = Invoke-Native -Program "git" -Arguments @("-C", $RepoRoot, "tag", "-a", $tag, $sha, "-F", $messagePath)
        if ($tagged.Code -ne 0) { Stop-Release $step "git tag failed: $($tagged.Output -join ' ')" }
        if ((& git -C $RepoRoot cat-file -t "refs/tags/$tag").Trim() -ne "tag" -or
            (& git -C $RepoRoot rev-parse "refs/tags/$tag^{commit}").Trim() -ne $sha) {
            Stop-Release $step "The local tag $tag is not an annotated tag of $sha; it was not pushed."
        }
        $pushed = Invoke-Native -Program "git" -Arguments @("-C", $RepoRoot, "push", "origin", "refs/tags/$tag")
        if ($pushed.Code -ne 0) { Stop-Release $step "git push of $tag failed: $($pushed.Output -join ' ')" }
        Write-Session "Pushed the annotated tag $tag"
        $assets = @((Join-Path $assetsDir $addinName), (Join-Path $assetsDir $demoName), $pdfPath,
            (Join-Path $assetsDir "release-evidence.json"))
        $created = Invoke-Native -Program "gh" -Arguments (@("release", "create", $tag, "--repo", $Repository,
                "--verify-tag", "--title", "$tag $([char]0x2014) $Title", "--notes-file", $notesPath) + $assets)
        foreach ($line in $created.Output) { Write-Session "  $line" }
        if ($created.Code -ne 0) {
            Stop-Release $step "gh release create failed. The tag $tag is pushed; create the release by hand from $assetsDir."
        }
        $script:summary.published = $true
        $download = Invoke-Native -Program "gh" -Arguments @("release", "download", $tag, "--repo", $Repository,
            "--dir", $folders.published)
        if ($download.Code -ne 0) { Stop-Release $step "The published assets could not be downloaded again." }
        foreach ($asset in $assets) {
            $name = [System.IO.Path]::GetFileName($asset)
            if ((Get-Sha256 (Join-Path $folders.published $name)) -ne (Get-Sha256 $asset)) {
                Stop-Release $step "The published $name differs from the tested file."
            }
        }
        $check = Invoke-Python -Python $python -Arguments @((Join-Path $RepoRoot "tools\check_release.py"),
            "--root", $RepoRoot, "--evidence", (Join-Path $folders.published "release-evidence.json"),
            "--assets-dir", $folders.published, "--static-report", (Join-Path $folders.static "static-checks.json"),
            "--require-tag", "--output", (Join-Path $releaseDir "published-check.json"))
        if ($check.Code -ne 0) { Stop-Release $step "check_release.py --require-tag refused the published release." }
        Complete-Step $step "$tag published; downloads match the tested files and validate with the tag"
        $script:summary.result = "PASS"
    }
}
catch [System.OperationCanceledException] {
    $script:summary.reason = $_.Exception.Message
    Write-Session $_.Exception.Message
}
catch {
    $script:summary.result = "FAIL"
    $script:summary.reason = $_.Exception.Message
    Write-Session ("Stopped: " + $_.Exception.Message)
}

# ----------------------------------------------------------------------------
# Record
# ----------------------------------------------------------------------------
$script:summary.finished_at = Get-Timestamp
[System.IO.File]::WriteAllText((Join-Path $releaseDir "release.json"),
    ($script:summary | ConvertTo-Json -Depth 6), $utf8NoBom)
$report = @(
    $(if ($Rehearse) { "Release rehearsal (steps 2 to 7; not release evidence)" } elseif ($DryRun) { "Release dry run" } else { "Release" }),
    "",
    "Result:     $($script:summary.result)",
    "Tag:        $tag$(if ($script:summary.published) { ' (published)' } elseif ($DryRun) { ' (dry run, not tagged)' } else { ' (not published)' })",
    "Candidate:  $sha",
    "Title:      $Title",
    "Static run: $($script:summary.static_run)",
    $(if ($script:summary.resumed) { "Resumed:    kept steps $($script:summary.kept_steps -join ', ')$(if ($script:resumedFrom) { ' (run at ' + $script:resumedFrom + ')' } else { '' })" } else { "Resumed:    no" }),
    "",
    "Steps:")
foreach ($item in $script:summary.steps) {
    $report += "  $($item.status.PadRight(7)) $($item.name)$(if ($item.detail) { ' - ' + $item.detail } else { '' })"
}
if ($script:summary.reason) {
    $report += @("", "Stopped because: $($script:summary.reason)")
}
$report += @("", "Started:  $($script:summary.started_at)", "Finished: $($script:summary.finished_at)",
    "Folder:   $releaseDir")
[System.IO.File]::WriteAllLines((Join-Path $releaseDir "release.txt"), [string[]]$report, $utf8NoBom)
Write-Host ""
foreach ($line in $report) {
    Write-Host $line
}
if ($script:summary.result -eq "PASS") {
    exit 0
}
exit 1
