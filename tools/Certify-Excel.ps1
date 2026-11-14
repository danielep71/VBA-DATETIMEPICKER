<#
.SYNOPSIS
Build, compile and run the regression pack of the checked-out commit in Excel
with one command, and write an evidence folder that names the commit and host.

.DESCRIPTION
For the commit checked out in this repository:
1. reads the full SHA and refuses a working tree with changes to tracked files,
   unless -AllowDirty is given (the evidence then says it is not
   certification-grade);
2. creates the evidence folder outside the checkout and starts session.log;
3. records Windows, Excel version and build, Office bitness, the macro setting
   and "Trust access to the VBA project object model", as found. It never
   changes any of them;
4. starts a visible Excel instance of its own, adds a new workbook and imports
   the components of one configuration of tools\vba-projects.json
   (packaged-hosts by default) in their listed order. Each component must keep
   its repository name;
5. adds a generated CertificationRunner module that calls the harness entry
   point under an error handler, so a raised error is recorded instead of
   opening Visual Basic's modal "Run-time error" window. The module is not part
   of the candidate's source;
6. compiles the project with Debug > Compile VBAProject;
7. runs TST_DP_RunAll, or TST_DP_RunAll_WithUISmoke with -WithUISmoke;
8. reads the harness summary and every result row from the TST_DP_RESULTS
   worksheet, saves the workbook in the evidence folder, closes it and quits
   Excel;
9. writes summary.json, results.tsv and certification.txt, and opens the
   folder.

The exit code is 0 only when the import and compile passed, the harness ran
and reported State=PASS with no cleanup failure, and the working tree was
clean. Everything else exits 1, and the evidence says why.

This automates the compile and regression steps of RELEASING.md. It does not
replace the manual matrices recorded for a release certification.

.REQUIREMENTS
- Windows with desktop Excel, and Git on the PATH.
- "Trust access to the VBA project object model" already on (Excel > File >
  Options > Trust Center > Trust Center Settings > Macro Settings). The import
  needs it; this script never turns it on.
- Close other Excel work first: the harness changes application-wide state
  (shortcut, context menu, events) and restores it at the end.

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Certify-Excel.ps1

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Certify-Excel.ps1 -WithUISmoke
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$BundleDir,
    [Parameter(Mandatory = $false)]
    [string]$Configuration = "packaged-hosts",
    [Parameter(Mandatory = $false)]
    [string]$Operator = "Repository owner, local Windows workstation",
    [Parameter(Mandatory = $false)]
    [string]$RepoRoot,
    [switch]$WithUISmoke,
    [switch]$AllowDirty,
    [switch]$NoOpen
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $RepoRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
}
$RepoRoot = (Resolve-Path -LiteralPath $RepoRoot).Path

$compileControlId = 578
$runnerName = "CertificationRunner"
$harnessModule = "M_cDP_Test"
$resultSheetName = "TST_DP_RESULTS"
$summaryValueColumn = 11
$xlOpenXMLWorkbookMacroEnabled = 52
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$script:sessionLog = $null
$entryPoint = $(if ($WithUISmoke) { "TST_DP_RunAll_WithUISmoke" } else { "TST_DP_RunAll" })

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

function Get-RegistryValue {
    param([string]$Path, [string]$Name)
    $item = Get-ItemProperty -LiteralPath $Path -Name $Name -ErrorAction SilentlyContinue
    if ($null -eq $item) {
        return $null
    }
    return $item.$Name
}

function Get-SecurityValue {
    # A value set by group policy takes precedence over the user's own setting.
    param([string]$OfficeVersion, [string]$Name)
    foreach ($root in @("HKCU:\Software\Policies\Microsoft\Office", "HKCU:\Software\Microsoft\Office")) {
        $value = Get-RegistryValue -Path "$root\$OfficeVersion\Excel\Security" -Name $Name
        if ($null -ne $value) {
            return $value
        }
    }
    return $null
}

function Get-MacroPolicy {
    param([string]$OfficeVersion)
    switch (Get-SecurityValue -OfficeVersion $OfficeVersion -Name "VBAWarnings") {
        1 { return "Enable VBA macros" }
        2 { return "Disable VBA macros with notification" }
        3 { return "Disable VBA macros except digitally signed macros" }
        4 { return "Disable VBA macros without notification" }
        default { return "Disable VBA macros with notification (default; no setting found)" }
    }
}

function Get-VbaProjectAccess {
    param([string]$OfficeVersion)
    if ((Get-SecurityValue -OfficeVersion $OfficeVersion -Name "AccessVBOM") -eq 1) {
        return "On"
    }
    return "Off"
}

function Invoke-Macro {
    # Application.Run with the workbook-qualified name; arguments are passed through.
    param([object]$Excel, [string]$Qualified, [string]$Procedure, [object[]]$Arguments = @())
    switch ($Arguments.Count) {
        0 { return $Excel.Run($Qualified + $Procedure) }
        default { return $Excel.Run($Qualified + $Procedure, $Arguments[0]) }
    }
}

function ConvertTo-CellText {
    # One worksheet value as single-line text for a tab-separated file.
    param([object]$Value)
    if ($null -eq $Value) {
        return ""
    }
    return ([string]$Value) -replace "[`t`r`n]+", " "
}

# ----------------------------------------------------------------------------
# Candidate, configuration and evidence folder
# ----------------------------------------------------------------------------
$sha = (& git -C $RepoRoot rev-parse HEAD).Trim()
if ($LASTEXITCODE -ne 0 -or $sha -notmatch '^[0-9a-f]{40}$') {
    throw "Cannot read the checked-out commit of $RepoRoot. Is Git on the PATH?"
}
$branch = (& git -C $RepoRoot rev-parse --abbrev-ref HEAD).Trim()
$changes = & git -C $RepoRoot status --porcelain --untracked-files=no
if ($LASTEXITCODE -ne 0) {
    throw "git status failed in $RepoRoot."
}
$dirty = -not [string]::IsNullOrWhiteSpace(($changes | Out-String))
if ($dirty -and -not $AllowDirty) {
    throw ("The working tree has changes to tracked files. Run on a clean checkout " +
           "(git status shows none), or pass -AllowDirty to try the tool without certifying.")
}

$manifestPath = Join-Path $RepoRoot "tools\vba-projects.json"
$manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$configurationNode = $manifest.configurations.PSObject.Properties | Where-Object { $_.Name -eq $Configuration }
if ($null -eq $configurationNode) {
    throw "tools\vba-projects.json has no configuration named '$Configuration'."
}
$components = @()
foreach ($path in @($configurationNode.Value.modules)) {
    $components += [pscustomobject]@{
        path   = [string]$path
        module = [System.IO.Path]::GetFileNameWithoutExtension([string]$path)
        file   = Join-Path $RepoRoot (([string]$path) -replace '/', '\')
    }
}
if ($components.Count -eq 0) {
    throw "Configuration '$Configuration' lists no modules."
}
foreach ($component in $components) {
    if (-not (Test-Path -LiteralPath $component.file)) {
        throw "Configured component is missing from the checkout: $($component.path)"
    }
}
if (-not ($components | Where-Object { $_.module -eq $harnessModule })) {
    throw "Configuration '$Configuration' does not include the regression harness $harnessModule."
}

$stamp = (Get-Date).ToString("yyyyMMdd-HHmmss")
if ([string]::IsNullOrWhiteSpace($BundleDir)) {
    $BundleDir = Join-Path (Split-Path -Parent $RepoRoot) ("DATETIMEPICKER-excel-" + $sha.Substring(0, 12) + "-" + $stamp)
}
if (Test-Path -LiteralPath $BundleDir) {
    if (@(Get-ChildItem -LiteralPath $BundleDir -Force).Count -gt 0) {
        throw "The evidence folder is not empty: $BundleDir. Evidence is never overwritten."
    }
}
else {
    New-Item -ItemType Directory -Path $BundleDir | Out-Null
}
$BundleDir = (Resolve-Path -LiteralPath $BundleDir).Path
if ($BundleDir.StartsWith($RepoRoot + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "The evidence folder must be outside the checkout: $BundleDir"
}
$script:sessionLog = Join-Path $BundleDir "session.log"
$startedAt = Get-Timestamp
Write-Session "Excel regression run of $sha ($branch)"
Write-Session "Checkout $RepoRoot; evidence $BundleDir; configuration $Configuration; entry point $entryPoint"
if ($dirty) {
    Write-Session "The working tree has changes (-AllowDirty): this run is not certification-grade"
}

# ----------------------------------------------------------------------------
# Environment, as found
# ----------------------------------------------------------------------------
$osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
$osName = ($osInfo.Caption -replace '^Microsoft\s+', '').Trim()
$osText = "{0} (version {1}, build {2})" -f $osName, $osInfo.Version, $osInfo.BuildNumber
$clickToRun = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
$platform = Get-RegistryValue -Path $clickToRun -Name "Platform"
$officeBitness = switch ($platform) {
    "x64" { "64-bit" }
    "x86" { "32-bit" }
    default { "unknown" }
}
$versionToReport = Get-RegistryValue -Path $clickToRun -Name "VersionToReport"
Write-Session "Windows: $osText"

$summary = [ordered]@{
    tool               = "tools/Certify-Excel.ps1"
    operator           = $Operator
    candidate_sha      = $sha
    branch             = $branch
    working_tree_clean = (-not $dirty)
    configuration      = $Configuration
    entry_point        = $entryPoint
    started_at         = $startedAt
    finished_at        = $null
    host               = [ordered]@{
        os                 = $osText
        excel_version      = "unknown"
        excel_build        = "unknown"
        office_bitness     = $officeBitness
        macro_policy       = "unknown"
        vba_project_access = "unknown"
    }
    import             = [ordered]@{ status = "NOT_RUN"; components = @() }
    compile            = [ordered]@{ status = "NOT_RUN"; detail = "Not run" }
    harness            = [ordered]@{
        status           = "NOT_RUN"
        error            = $null
        summary_line     = $null
        state            = $null
        run              = $null
        passed           = $null
        failed           = $null
        cleanup_failures = $null
        result_rows      = 0
    }
    workbook           = $null
    result             = "FAIL"
    reasons            = @()
}

$workbookPath = Join-Path $BundleDir "DATETIMEPICKER-regression.xlsm"
$excel = $null
$workbook = $null
try {
    # ------------------------------------------------------------------------
    # Excel and its security settings, read and never changed
    # ------------------------------------------------------------------------
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $officeVersion = [string]$excel.Version
    $summary.host.excel_version = $officeVersion
    if ($null -ne $versionToReport) {
        $summary.host.excel_build = "Version $versionToReport (build $($excel.Build))"
    }
    else {
        $summary.host.excel_build = "Version $officeVersion (build $($excel.Build))"
    }
    $summary.host.macro_policy = Get-MacroPolicy -OfficeVersion $officeVersion
    $summary.host.vba_project_access = Get-VbaProjectAccess -OfficeVersion $officeVersion
    Write-Session "Excel: $($summary.host.excel_build), $officeBitness"
    Write-Session "Macro setting: $($summary.host.macro_policy); VBA project access: $($summary.host.vba_project_access)"
    if ($summary.host.vba_project_access -ne "On") {
        throw ("'Trust access to the VBA project object model' is off. The import needs it: " +
               "in Excel choose File > Options > Trust Center > Trust Center Settings > Macro Settings " +
               "and tick it, if your policy allows. This script never changes it.")
    }

    # ------------------------------------------------------------------------
    # Import the exact source into a new workbook
    # ------------------------------------------------------------------------
    $workbook = $excel.Workbooks.Add()
    $qualified = "'" + $workbook.Name + "'!"
    $project = $workbook.VBProject
    $imported = @()
    foreach ($component in $components) {
        $added = $project.VBComponents.Import($component.file)
        $name = [string]$added.Name
        if ($name -ne $component.module) {
            $summary.import = [ordered]@{ status = "FAIL"; components = $imported }
            throw "The component $($component.path) was imported as $name, not $($component.module)."
        }
        $imported += $component.path
        Write-Session "Imported $($component.path) as $name"
    }
    $summary.import = [ordered]@{ status = "PASS"; components = $imported }

    # ------------------------------------------------------------------------
    # Add the generated runner
    # ------------------------------------------------------------------------
    # Called directly with Application.Run, an error raised by the harness would
    # open Visual Basic's modal "Run-time error" window and hold the run until
    # someone pressed End. The runner calls the entry point under an error
    # handler and returns RAN or the error, so the run records it and goes on.
    foreach ($existing in $project.VBComponents) {
        if ([string]$existing.Name -eq $runnerName) {
            throw "The project already has a component named $runnerName."
        }
    }
    $runnerLines = @(
        "Option Explicit",
        "",
        "Public Function Ready() As String",
        "    Ready = ""READY""",
        "End Function",
        "",
        "Public Function RunHarness() As String",
        "    On Error GoTo Failed",
        "    $harnessModule.$entryPoint",
        "    RunHarness = ""RAN""",
        "    Exit Function",
        "Failed:",
        "    RunHarness = ""ERROR "" & CStr(Err.Number) & "": "" & Err.Description",
        "End Function",
        "")
    $runner = $project.VBComponents.Add(1)
    $runner.Name = $runnerName
    # "Require Variable Declaration" may already have inserted Option Explicit.
    if ($runner.CodeModule.CountOfLines -gt 0) {
        $runner.CodeModule.DeleteLines(1, $runner.CodeModule.CountOfLines)
    }
    $runner.CodeModule.AddFromString(($runnerLines -join "`r`n"))
    Write-Session "Added the generated $runnerName module; it is not part of the candidate's source"

    # ------------------------------------------------------------------------
    # Compile
    # ------------------------------------------------------------------------
    # Debug > Compile VBAProject acts on the VBE's active project and is greyed
    # out once that project is compiled. The VBE window is shown so that the
    # command's state is current, and the state is read afresh.
    $vbe = $excel.VBE
    $vbe.MainWindow.Visible = $true
    try {
        $vbe.ActiveVBProject = $project
    }
    catch {
        Write-Session ("The project could not be made the active one: " + $_.Exception.Message)
    }
    $compile = $vbe.CommandBars.FindControl(1, $compileControlId)
    if ($null -eq $compile) {
        $summary.compile = [ordered]@{ status = "FAIL"; detail = "Debug > Compile VBAProject was not found in the VBE" }
        throw "Cannot find Debug > Compile VBAProject."
    }
    if ($compile.Enabled) {
        $compile.Execute()
    }
    $compiled = $false
    for ($attempt = 0; $attempt -lt 10 -and -not $compiled; $attempt++) {
        Start-Sleep -Milliseconds 500
        $compiled = -not $vbe.CommandBars.FindControl(1, $compileControlId).Enabled
    }
    $compileDetail = "Debug > Compile VBAProject completed with no error"
    if (-not $compiled) {
        # The command can stay enabled without a compile error, so the operator
        # looks at the VBE instead of the run stopping on a guess.
        Write-Host ""
        Write-Host "Debug > Compile VBAProject still looks enabled. In the VBE window, choose"
        Write-Host "Debug > Compile VBAProject yourself. Type Y if it completes with no error and"
        Write-Host "the command is then greyed out; type N if it reports a compile error."
        $answer = Read-Host "Compiled with no error? (Y/N)"
        if ($answer -notmatch '^\s*[yY]') {
            $summary.compile = [ordered]@{ status = "FAIL"; detail = "Debug > Compile VBAProject reported an error, as the operator saw" }
            throw "The project does not compile; see the VBE."
        }
        $compileDetail = "Debug > Compile VBAProject completed with no error, as the operator saw in the VBE"
    }
    $summary.compile = [ordered]@{ status = "PASS"; detail = $compileDetail }
    Write-Session $compileDetail
    $vbe.MainWindow.Visible = $false

    $ready = [string](Invoke-Macro -Excel $excel -Qualified $qualified -Procedure "$runnerName.Ready")
    if ($ready -ne "READY") {
        throw "$runnerName.Ready answered '$ready', not READY."
    }
    Write-Session "The project's code runs"

    # ------------------------------------------------------------------------
    # Run the regression harness
    # ------------------------------------------------------------------------
    Write-Session "Running $harnessModule.$entryPoint; do not use Excel until it returns"
    try {
        $answer = [string](Invoke-Macro -Excel $excel -Qualified $qualified -Procedure "$runnerName.RunHarness")
        if ($answer -eq "RAN") {
            $summary.harness.status = "RAN"
        }
        else {
            $summary.harness.status = "ERROR"
            $summary.harness.error = $(if ($answer.StartsWith("ERROR ")) { $answer.Substring(6) } else { "the runner answered '$answer'" })
        }
    }
    catch {
        $summary.harness.status = "ERROR"
        $summary.harness.error = $_.Exception.Message
    }
    Write-Session ("$entryPoint " + $summary.harness.status + $(if ($summary.harness.error) { ": " + $summary.harness.error } else { "" }))

    # ------------------------------------------------------------------------
    # Read the harness's own record
    # ------------------------------------------------------------------------
    $sheet = $null
    foreach ($candidate in $workbook.Worksheets) {
        if ([string]$candidate.Name -eq $resultSheetName) {
            $sheet = $candidate
        }
    }
    if ($null -eq $sheet) {
        Write-Session "The harness wrote no $resultSheetName worksheet"
    }
    else {
        $summary.harness.run = [string]$sheet.Cells.Item(5, $summaryValueColumn).Text
        $summary.harness.passed = [string]$sheet.Cells.Item(6, $summaryValueColumn).Text
        $summary.harness.failed = [string]$sheet.Cells.Item(7, $summaryValueColumn).Text
        $summary.harness.state = [string]$sheet.Cells.Item(8, $summaryValueColumn).Text
        $values = $sheet.UsedRange.Value2
        $rows = New-Object System.Collections.Generic.List[string]
        if ($values -is [array] -and $values.Rank -eq 2) {
            for ($r = $values.GetLowerBound(0); $r -le $values.GetUpperBound(0); $r++) {
                $cells = @()
                for ($c = $values.GetLowerBound(1); $c -le $values.GetUpperBound(1); $c++) {
                    $text = ConvertTo-CellText $values[$r, $c]
                    $cells += $text
                    if ($null -eq $summary.harness.summary_line -and
                        $text.StartsWith("State=") -and $text.Contains("CleanupFailures=")) {
                        $summary.harness.summary_line = $text
                    }
                }
                $rows.Add(($cells -join "`t"))
            }
        }
        $summary.harness.result_rows = $rows.Count
        [System.IO.File]::WriteAllLines((Join-Path $BundleDir "results.tsv"), $rows, $utf8NoBom)
        if ($null -ne $summary.harness.summary_line -and
            $summary.harness.summary_line -match 'CleanupFailures=(\d+)') {
            $summary.harness.cleanup_failures = [int]$Matches[1]
        }
        Write-Session ("Harness summary: " + $(if ($summary.harness.summary_line) { $summary.harness.summary_line } else {
                    "State=$($summary.harness.state); Run=$($summary.harness.run); Passed=$($summary.harness.passed); Failed=$($summary.harness.failed)" }))
    }

    # ------------------------------------------------------------------------
    # Keep the workbook for inspection
    # ------------------------------------------------------------------------
    try {
        $workbook.SaveAs($workbookPath, $xlOpenXMLWorkbookMacroEnabled)
        $summary.workbook = [System.IO.Path]::GetFileName($workbookPath)
        Write-Session "Workbook saved as $workbookPath"
    }
    catch {
        # Evidence without its workbook is incomplete, so it cannot pass.
        $summary.reasons += ("The workbook could not be saved: " + $_.Exception.Message)
        Write-Session ("The workbook could not be saved: " + $_.Exception.Message)
    }
}
catch {
    $summary.reasons += ("Stopped: " + $_.Exception.Message)
    Write-Session ("Stopped: " + $_.Exception.Message)
}
finally {
    if ($null -ne $workbook) {
        try {
            $workbook.Close($false)
            Write-Session "Workbook closed"
        }
        catch {
            Write-Session ("The workbook could not be closed: " + $_.Exception.Message)
        }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }
    if ($null -ne $excel) {
        try {
            $excel.Quit()
            Write-Session "Excel closed"
        }
        catch {
            Write-Session ("Excel could not be closed: " + $_.Exception.Message)
        }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# ----------------------------------------------------------------------------
# Verdict and evidence
# ----------------------------------------------------------------------------
if ($dirty) { $summary.reasons += "The working tree had changes to tracked files" }
if ($summary.import.status -ne "PASS") { $summary.reasons += "Import: $($summary.import.status)" }
if ($summary.compile.status -ne "PASS") { $summary.reasons += "Compile: $($summary.compile.status)" }
if ($summary.harness.status -ne "RAN") { $summary.reasons += "Harness: $($summary.harness.status)" }
if ($summary.harness.state -ne "PASS") { $summary.reasons += "Harness state: $($summary.harness.state)" }
if ($null -eq $summary.harness.cleanup_failures) {
    $summary.reasons += "The harness summary line (State=...; CleanupFailures=...) was not found"
}
elseif ($summary.harness.cleanup_failures -ne 0) {
    $summary.reasons += "Cleanup failures: $($summary.harness.cleanup_failures)"
}
if ($summary.reasons.Count -eq 0) {
    $summary.result = "PASS"
}
$summary.finished_at = Get-Timestamp
[System.IO.File]::WriteAllText((Join-Path $BundleDir "summary.json"),
    ($summary | ConvertTo-Json -Depth 6), $utf8NoBom)

$report = @(
    "Excel regression run",
    "",
    "Result:          $($summary.result)",
    "Candidate:       $sha ($branch)",
    "Working tree:    $(if ($dirty) { 'CHANGED (not certification-grade)' } else { 'clean' })",
    "Configuration:   $Configuration",
    "Entry point:     $harnessModule.$entryPoint",
    "Excel:           $($summary.host.excel_build), $officeBitness",
    "Windows:         $osText",
    "Macro setting:   $($summary.host.macro_policy)",
    "Import:          $($summary.import.status)",
    "Compile:         $($summary.compile.status) - $($summary.compile.detail)",
    "Harness:         $($summary.harness.status)$(if ($summary.harness.error) { ' - ' + $summary.harness.error } else { '' })",
    "",
    "Harness summary line, verbatim:",
    "$(if ($summary.harness.summary_line) { $summary.harness.summary_line } else { '(not found)' })",
    "",
    "Started:  $startedAt",
    "Finished: $($summary.finished_at)")
if ($summary.reasons.Count -gt 0) {
    $report += @("", "Not passing because:")
    foreach ($reason in $summary.reasons) {
        $report += "- $reason"
    }
}
[System.IO.File]::WriteAllLines((Join-Path $BundleDir "certification.txt"), $report, $utf8NoBom)

Write-Host ""
foreach ($line in $report) {
    Write-Host $line
}
Write-Host ""
Write-Host "Evidence: $BundleDir"
if (-not $NoOpen) {
    try {
        Invoke-Item -LiteralPath $BundleDir
    }
    catch {
    }
}
if ($summary.result -eq "PASS") {
    exit 0
}
exit 1
