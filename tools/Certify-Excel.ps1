<#
.SYNOPSIS
Run the regression pack in Excel with one command, either on the checked-out
source or inside a built package, and write an evidence folder that names the
commit, the package and the host.

.DESCRIPTION
Source run (the default), for the commit checked out in this repository:
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
8. reads the harness summary, the suites dispatched and completed, and every
   result row from the TST_DP_RESULTS worksheet, saves the workbook in the
   evidence folder, closes it and quits Excel. Each close step is tried even
   when an earlier one fails, and each failure is recorded
   (tools\ExcelCleanup.psm1);
9. writes summary.json, results.tsv and certification.txt, and opens the
   folder.

Package run (-Package <file>), for a built .xlam or demo .xlsm (#222):
1. and 2. as above; the package's size and SHA-256 are recorded first;
3. runs two sessions, each in a new Excel instance: TST_DP_RunAll, then
   TST_DP_RunAll_WithUISmoke. Each session sets that instance's
   Application.AutomationSecurity to Low, so the package's macros run
   without a "Security Notice" (the property is not saved and the Trust
   Center settings are unchanged), then opens the package read-only (an
   add-in also gets a blank workbook to hold the result sheet), adds the
   generated runner to the open package in memory, compiles, runs the pack,
   reads the results and closes everything without saving;
4. after each session, checks that the package's SHA-256 has not changed;
5. writes summary.json, results-standard.tsv, results-ui_smoke.tsv and
   certification.txt.
The package run proves what these exact bytes did in Excel. Which commit built
them is recorded by the build evidence (tools\Build-Addin.cmd), not here.

The exit code is 0 only when every compile passed, the harness ran and
reported State=PASS with no cleanup failure, every dispatched suite completed,
Excel accepted every close request, the working tree was clean and, for a
package, its hash never changed. Everything else exits 1, and the evidence
says why.

This automates the compile and regression steps of RELEASING.md. It does not
replace the manual matrices recorded for a release certification.

.REQUIREMENTS
- Windows with desktop Excel, and Git on the PATH.
- "Trust access to the VBA project object model" already on (Excel > File >
  Options > Trust Center > Trust Center Settings > Macro Settings). The import
  and the runner need it; this script never turns it on.
- Close other Excel work first: the harness changes application-wide state
  (shortcut, context menu, events) and restores it at the end.

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Certify-Excel.ps1

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Certify-Excel.ps1 -WithUISmoke

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Certify-Excel.ps1 -Package ..\DATETIMEPICKER-addin-...\DATETIMEPICKER.v1.2.3.xlam
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$BundleDir,
    [Parameter(Mandatory = $false)]
    [string]$Configuration = "packaged-hosts",
    [Parameter(Mandatory = $false)]
    [string]$Package,
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
Import-Module (Join-Path $PSScriptRoot "ExcelCleanup.psm1") -Force

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
$msoAutomationSecurityLow = 1
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$script:sessionLog = $null
$script:cleanupFailures = New-Object System.Collections.Generic.List[string]
$packageMode = -not [string]::IsNullOrWhiteSpace($Package)
$entryPoint = $(if ($WithUISmoke) { "TST_DP_RunAll_WithUISmoke" } else { "TST_DP_RunAll" })
$packs = [ordered]@{
    standard = "TST_DP_RunAll"
    ui_smoke = "TST_DP_RunAll_WithUISmoke"
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

function ConvertTo-Count {
    # A summary cell as a whole number, or $null when it holds none.
    param([object]$Value)
    $number = 0
    if ([int]::TryParse(([string]$Value).Trim(), [ref]$number)) {
        return $number
    }
    return $null
}

function Get-Sha256 {
    param([string]$Path)
    return (Get-FileHash -Algorithm SHA256 -LiteralPath $Path).Hash.ToLowerInvariant()
}

function Close-Excel {
    # Every step is tried and every failure kept for the verdict: Excel may
    # still be running, or still hold the package open.
    param([object]$Excel, [object[]]$Workbooks, [string]$Label)
    $logCleanup = { param([string]$Text) Write-Session $Text }
    foreach ($failure in @(Close-ExcelInstance -Excel $Excel -Workbooks $Workbooks -Label $Label -Log $logCleanup)) {
        $script:cleanupFailures.Add($failure)
    }
}

function Enable-PackageMacros {
    # A package opened by a script runs its macros only as this instance's
    # Application.AutomationSecurity allows. With the Trust Center's
    # "Disable VBA macros with notification", Excel would otherwise stop at a
    # "Security Notice" for every pack until someone chose Enable Macros. The
    # property belongs to this Excel instance only and is not saved: the Trust
    # Center settings stay as they are. A policy may refuse it; then Excel asks,
    # and the run waits for the answer.
    param([object]$Excel)
    try {
        $Excel.AutomationSecurity = $msoAutomationSecurityLow
    }
    catch {
        Write-Session ("AutomationSecurity could not be set: " + $_.Exception.Message)
    }
    $level = [int]$Excel.AutomationSecurity
    if ($level -eq $msoAutomationSecurityLow) {
        $summary.host.package_macros = "Enabled for files this Excel instance opens (AutomationSecurity=1)"
    }
    else {
        $summary.host.package_macros = "AutomationSecurity=$level; Excel asks before running the package's macros"
        Write-Host ""
        Write-Host "Excel will ask whether to enable the package's macros. Choose Enable Macros;"
        Write-Host "the run waits until you do."
    }
    Write-Session "Package macros: $($summary.host.package_macros)"
}

function New-HarnessRecord {
    return [ordered]@{
        status            = "NOT_RUN"
        error             = $null
        summary_line      = $null
        state             = $null
        run               = $null
        passed            = $null
        failed            = $null
        cleanup_failures  = $null
        suites_dispatched = $null
        suites_completed  = $null
        result_rows       = 0
        results_file      = $null
    }
}

function Start-Excel {
    # A visible instance of our own. The host is recorded from the first one;
    # the run stops when the VBA project cannot be reached.
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    if ($summary.host.excel_version -eq "unknown") {
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
    }
    return $excel
}

function Assert-VbaProjectAccess {
    if ($summary.host.vba_project_access -ne "On") {
        throw ("'Trust access to the VBA project object model' is off. The runner needs it: " +
               "in Excel choose File > Options > Trust Center > Trust Center Settings > Macro Settings " +
               "and tick it, if your policy allows. This script never changes it.")
    }
}

function Add-Runner {
    # Called directly with Application.Run, an error raised by the harness
    # would open Visual Basic's modal "Run-time error" window and hold the run
    # until someone pressed End. The runner calls the entry point under an
    # error handler and returns RAN or the error, so the run records it and
    # goes on. It is never saved: a package is open read-only and closed
    # without saving.
    param([object]$Project, [string]$EntryPoint)
    foreach ($existing in $Project.VBComponents) {
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
        "    $harnessModule.$EntryPoint",
        "    RunHarness = ""RAN""",
        "    Exit Function",
        "Failed:",
        "    RunHarness = ""ERROR "" & CStr(Err.Number) & "": "" & Err.Description",
        "End Function",
        "")
    $runner = $Project.VBComponents.Add(1)
    $runner.Name = $runnerName
    # "Require Variable Declaration" may already have inserted Option Explicit.
    if ($runner.CodeModule.CountOfLines -gt 0) {
        $runner.CodeModule.DeleteLines(1, $runner.CodeModule.CountOfLines)
    }
    $runner.CodeModule.AddFromString(($runnerLines -join "`r`n"))
    Write-Session "Added the generated $runnerName module; it is not part of the candidate's source"
}

function Invoke-Compile {
    # Debug > Compile VBAProject acts on the VBE's active project and is greyed
    # out once that project is compiled. The VBE window is shown so that the
    # command's state is current, and the state is read afresh.
    param([object]$Excel, [object]$Project)
    $vbe = $Excel.VBE
    $vbe.MainWindow.Visible = $true
    try {
        $vbe.ActiveVBProject = $Project
    }
    catch {
        Write-Session ("The project could not be made the active one: " + $_.Exception.Message)
    }
    $compile = $vbe.CommandBars.FindControl(1, $compileControlId)
    if ($null -eq $compile) {
        return [ordered]@{ status = "FAIL"; detail = "Debug > Compile VBAProject was not found in the VBE" }
    }
    if ($compile.Enabled) {
        $compile.Execute()
    }
    $compiled = $false
    for ($attempt = 0; $attempt -lt 10 -and -not $compiled; $attempt++) {
        Start-Sleep -Milliseconds 500
        $compiled = -not $vbe.CommandBars.FindControl(1, $compileControlId).Enabled
    }
    $detail = "Debug > Compile VBAProject completed with no error"
    if (-not $compiled) {
        # The command can stay enabled without a compile error, so the operator
        # looks at the VBE instead of the run stopping on a guess.
        Write-Host ""
        Write-Host "Debug > Compile VBAProject still looks enabled. In the VBE window, choose"
        Write-Host "Debug > Compile VBAProject yourself. Type Y if it completes with no error and"
        Write-Host "the command is then greyed out; type N if it reports a compile error."
        $answer = Read-Host "Compiled with no error? (Y/N)"
        if ($answer -notmatch '^\s*[yY]') {
            return [ordered]@{ status = "FAIL"; detail = "Debug > Compile VBAProject reported an error, as the operator saw" }
        }
        $detail = "Debug > Compile VBAProject completed with no error, as the operator saw in the VBE"
    }
    $vbe.MainWindow.Visible = $false
    Write-Session $detail
    return [ordered]@{ status = "PASS"; detail = $detail }
}

function Invoke-Harness {
    # Checks that the project's code runs, then runs the entry point through
    # the runner and records RAN or the error.
    param([object]$Excel, [string]$Qualified, [string]$EntryPoint, [System.Collections.IDictionary]$Record)
    $ready = [string](Invoke-Macro -Excel $Excel -Qualified $Qualified -Procedure "$runnerName.Ready")
    if ($ready -ne "READY") {
        throw "$runnerName.Ready answered '$ready', not READY."
    }
    Write-Session "The project's code runs"
    Write-Session "Running $harnessModule.$EntryPoint; do not use Excel until it returns"
    try {
        $answer = [string](Invoke-Macro -Excel $Excel -Qualified $Qualified -Procedure "$runnerName.RunHarness")
        if ($answer -eq "RAN") {
            $Record.status = "RAN"
        }
        else {
            $Record.status = "ERROR"
            $Record.error = $(if ($answer.StartsWith("ERROR ")) { $answer.Substring(6) } else { "the runner answered '$answer'" })
        }
    }
    catch {
        $Record.status = "ERROR"
        $Record.error = $_.Exception.Message
    }
    Write-Session ("$EntryPoint " + $Record.status + $(if ($Record.error) { ": " + $Record.error } else { "" }))
}

function Read-Harness {
    # Reads the harness's own record: the summary cells K5:K10, the summary
    # line and every result row. The result sheet is looked for in every open
    # workbook, because an add-in writes it into the active workbook.
    param([object]$Excel, [System.Collections.IDictionary]$Record, [string]$ResultsFile)
    $sheet = $null
    foreach ($book in $Excel.Workbooks) {
        foreach ($candidate in $book.Worksheets) {
            if ($null -eq $sheet -and [string]$candidate.Name -eq $resultSheetName) {
                $sheet = $candidate
            }
        }
    }
    if ($null -eq $sheet) {
        Write-Session "The harness wrote no $resultSheetName worksheet"
        return
    }
    $Record.run = ConvertTo-Count $sheet.Cells.Item(5, $summaryValueColumn).Text
    $Record.passed = ConvertTo-Count $sheet.Cells.Item(6, $summaryValueColumn).Text
    $Record.failed = ConvertTo-Count $sheet.Cells.Item(7, $summaryValueColumn).Text
    $Record.state = [string]$sheet.Cells.Item(8, $summaryValueColumn).Text
    $Record.suites_dispatched = ConvertTo-Count $sheet.Cells.Item(9, $summaryValueColumn).Text
    $Record.suites_completed = ConvertTo-Count $sheet.Cells.Item(10, $summaryValueColumn).Text
    $values = $sheet.UsedRange.Value2
    $rows = New-Object System.Collections.Generic.List[string]
    if ($values -is [array] -and $values.Rank -eq 2) {
        for ($r = $values.GetLowerBound(0); $r -le $values.GetUpperBound(0); $r++) {
            $cells = @()
            for ($c = $values.GetLowerBound(1); $c -le $values.GetUpperBound(1); $c++) {
                $text = ConvertTo-CellText $values[$r, $c]
                $cells += $text
                if ($null -eq $Record.summary_line -and
                    $text.StartsWith("State=") -and $text.Contains("CleanupFailures=")) {
                    $Record.summary_line = $text
                }
            }
            $rows.Add(($cells -join "`t"))
        }
    }
    $Record.result_rows = $rows.Count
    [System.IO.File]::WriteAllLines($ResultsFile, $rows, $utf8NoBom)
    $Record.results_file = [System.IO.Path]::GetFileName($ResultsFile)
    if ($null -ne $Record.summary_line -and $Record.summary_line -match 'CleanupFailures=(\d+)') {
        $Record.cleanup_failures = [int]$Matches[1]
    }
    Write-Session ("Harness summary: " + $(if ($Record.summary_line) { $Record.summary_line } else {
                "State=$($Record.state); Run=$($Record.run); Passed=$($Record.passed); Failed=$($Record.failed)" }))
    Write-Session "Suites: $($Record.suites_completed) completed of $($Record.suites_dispatched) dispatched"
}

function Get-HarnessReasons {
    # Why a harness record does not pass, one text per reason.
    param([System.Collections.IDictionary]$Record, [string]$Prefix)
    if ($Record.status -ne "RAN") { "${Prefix}Harness: $($Record.status)" }
    if ($Record.state -ne "PASS") { "${Prefix}Harness state: $($Record.state)" }
    if ($null -eq $Record.cleanup_failures) {
        "${Prefix}The harness summary line (State=...; CleanupFailures=...) was not found"
    }
    elseif ($Record.cleanup_failures -ne 0) {
        "${Prefix}Cleanup failures: $($Record.cleanup_failures)"
    }
    if ($null -eq $Record.suites_dispatched -or $null -eq $Record.suites_completed) {
        "${Prefix}The suite counts (K9:K10 of $resultSheetName) were not found"
    }
    elseif ($Record.suites_dispatched -eq 0 -or $Record.suites_completed -ne $Record.suites_dispatched) {
        "${Prefix}Suites completed: $($Record.suites_completed) of $($Record.suites_dispatched) dispatched"
    }
}

function Format-Suites {
    param([System.Collections.IDictionary]$Record)
    if ($null -eq $Record.suites_dispatched) {
        return "(not found)"
    }
    return "$($Record.suites_completed) completed of $($Record.suites_dispatched) dispatched"
}

# ----------------------------------------------------------------------------
# Candidate, configuration or package, and evidence folder
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

$components = @()
$packagePath = $null
$packageName = $null
$packageHost = $null
if ($packageMode) {
    if ($WithUISmoke) {
        throw "-WithUISmoke applies to a source run; a package run always runs both packs."
    }
    if (-not (Test-Path -LiteralPath $Package -PathType Leaf)) {
        throw "The package does not exist: $Package"
    }
    $packagePath = (Resolve-Path -LiteralPath $Package).Path
    $packageName = [System.IO.Path]::GetFileName($packagePath)
    $packageHost = switch ([System.IO.Path]::GetExtension($packagePath).ToLowerInvariant()) {
        ".xlam" { "xlam" }
        ".xlsm" { "xlsm" }
        default { throw "A package must be an .xlam add-in or an .xlsm workbook: $packageName" }
    }
}
else {
    $manifestPath = Join-Path $RepoRoot "tools\vba-projects.json"
    $manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $configurationNode = $manifest.configurations.PSObject.Properties | Where-Object { $_.Name -eq $Configuration }
    if ($null -eq $configurationNode) {
        throw "tools\vba-projects.json has no configuration named '$Configuration'."
    }
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
}

$stamp = (Get-Date).ToString("yyyyMMdd-HHmmss")
if ([string]::IsNullOrWhiteSpace($BundleDir)) {
    $prefix = $(if ($packageMode) { "DATETIMEPICKER-package-" } else { "DATETIMEPICKER-excel-" })
    $BundleDir = Join-Path (Split-Path -Parent $RepoRoot) ($prefix + $sha.Substring(0, 12) + "-" + $stamp)
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
if ($packageMode) {
    Write-Session "Excel regression run of the package $packageName, with the checkout at $sha ($branch)"
    Write-Session "Package $packagePath; evidence $BundleDir; packs $($packs.Values -join ', ')"
}
else {
    Write-Session "Excel regression run of $sha ($branch)"
    Write-Session "Checkout $RepoRoot; evidence $BundleDir; configuration $Configuration; entry point $entryPoint"
}
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
    mode               = $(if ($packageMode) { "package" } else { "source" })
    operator           = $Operator
    candidate_sha      = $sha
    branch             = $branch
    working_tree_clean = (-not $dirty)
    configuration      = $(if ($packageMode) { $null } else { $Configuration })
    entry_point        = $(if ($packageMode) { $null } else { $entryPoint })
    started_at         = $startedAt
    finished_at        = $null
    host               = [ordered]@{
        os                 = $osText
        excel_version      = "unknown"
        excel_build        = "unknown"
        office_bitness     = $officeBitness
        macro_policy       = "unknown"
        vba_project_access = "unknown"
        package_macros     = $null
    }
    import             = [ordered]@{ status = "NOT_RUN"; components = @() }
    compile            = [ordered]@{ status = "NOT_RUN"; detail = "Not run" }
    harness            = New-HarnessRecord
    workbook           = $null
    package            = $null
    packs              = $null
    cleanup            = [ordered]@{ status = "NOT_RUN"; failures = @() }
    result             = "FAIL"
    reasons            = @()
}
if ($packageMode) {
    # A package run has no import and no single harness record: each pack has
    # its own compile and harness record.
    $summary.import = $null
    $summary.compile = $null
    $summary.harness = $null
    $summary.package = [ordered]@{
        file   = $packageName
        host   = $packageHost
        bytes  = (Get-Item -LiteralPath $packagePath).Length
        sha256 = Get-Sha256 -Path $packagePath
    }
    $summary.packs = [ordered]@{}
    foreach ($key in $packs.Keys) {
        $summary.packs[$key] = [ordered]@{
            entry_point  = $packs[$key]
            compile      = [ordered]@{ status = "NOT_RUN"; detail = "Not run" }
            harness      = New-HarnessRecord
            sha256_after = $null
        }
    }
    Write-Session "Package $packageName ($packageHost), $($summary.package.bytes) bytes, SHA-256 $($summary.package.sha256)"
}

$workbookPath = Join-Path $BundleDir "DATETIMEPICKER-regression.xlsm"
$excel = $null
$workbook = $null
$hostBook = $null
try {
    if ($packageMode) {
        # --------------------------------------------------------------------
        # One fresh Excel instance per pack, the package open read-only
        # --------------------------------------------------------------------
        foreach ($key in $packs.Keys) {
            $pack = $summary.packs[$key]
            Write-Session "Pack $key ($($pack.entry_point)) in a new Excel instance"
            try {
                $excel = Start-Excel
                Assert-VbaProjectAccess
                Enable-PackageMacros -Excel $excel
                $workbook = $excel.Workbooks.Open($packagePath, 0, $true)
                if ($packageHost -eq "xlam") {
                    # An add-in has no window; the harness writes its result
                    # sheet into the active workbook.
                    $hostBook = $excel.Workbooks.Add()
                }
                Write-Session "Opened $packageName read-only$(if ($null -ne $hostBook) { ', with a blank workbook for the results' } else { '' })"
                $qualified = "'" + $workbook.Name + "'!"
                Add-Runner -Project $workbook.VBProject -EntryPoint $pack.entry_point
                $pack.compile = Invoke-Compile -Excel $excel -Project $workbook.VBProject
                if ($pack.compile.status -ne "PASS") {
                    throw "The package does not compile; see the VBE."
                }
                Invoke-Harness -Excel $excel -Qualified $qualified -EntryPoint $pack.entry_point -Record $pack.harness
                Read-Harness -Excel $excel -Record $pack.harness -ResultsFile (Join-Path $BundleDir "results-$key.tsv")
            }
            finally {
                Close-Excel -Excel $excel -Workbooks @($hostBook, $workbook) -Label "Pack $key"
                $hostBook = $null
                $workbook = $null
                $excel = $null
            }
            $pack.sha256_after = Get-Sha256 -Path $packagePath
            Write-Session "SHA-256 after the $key pack: $($pack.sha256_after)"
            if ($pack.sha256_after -ne $summary.package.sha256) {
                throw "The package changed on disk during the $key pack."
            }
        }
    }
    else {
        # --------------------------------------------------------------------
        # Import the exact source into a new workbook
        # --------------------------------------------------------------------
        $excel = Start-Excel
        Assert-VbaProjectAccess
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

        Add-Runner -Project $project -EntryPoint $entryPoint
        $summary.compile = Invoke-Compile -Excel $excel -Project $project
        if ($summary.compile.status -ne "PASS") {
            throw "The project does not compile; see the VBE."
        }
        Invoke-Harness -Excel $excel -Qualified $qualified -EntryPoint $entryPoint -Record $summary.harness
        Read-Harness -Excel $excel -Record $summary.harness -ResultsFile (Join-Path $BundleDir "results.tsv")

        # --------------------------------------------------------------------
        # Keep the workbook for inspection
        # --------------------------------------------------------------------
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
}
catch {
    $summary.reasons += ("Stopped: " + $_.Exception.Message)
    Write-Session ("Stopped: " + $_.Exception.Message)
}
finally {
    # Every step is tried and every failure kept: Excel may still be running.
    if ($null -ne $excel) {
        Close-Excel -Excel $excel -Workbooks @($hostBook, $workbook) -Label "Run"
    }
    $hostBook = $null
    $workbook = $null
    $excel = $null
}

# ----------------------------------------------------------------------------
# Verdict and evidence
# ----------------------------------------------------------------------------
Add-CleanupResult -Summary $summary -Failures $script:cleanupFailures.ToArray()
if ($dirty) { $summary.reasons += "The working tree had changes to tracked files" }
if ($packageMode) {
    foreach ($key in $packs.Keys) {
        $pack = $summary.packs[$key]
        if ($pack.compile.status -ne "PASS") { $summary.reasons += "Pack ${key}: compile $($pack.compile.status)" }
        $summary.reasons += @(Get-HarnessReasons -Record $pack.harness -Prefix "Pack ${key}: ")
        if ($pack.sha256_after -ne $summary.package.sha256) {
            $summary.reasons += "Pack ${key}: the package hash changed, or was not read after the pack"
        }
    }
}
else {
    if ($summary.import.status -ne "PASS") { $summary.reasons += "Import: $($summary.import.status)" }
    if ($summary.compile.status -ne "PASS") { $summary.reasons += "Compile: $($summary.compile.status)" }
    $summary.reasons += @(Get-HarnessReasons -Record $summary.harness -Prefix "")
}
if ($summary.reasons.Count -eq 0) {
    $summary.result = "PASS"
}
$summary.finished_at = Get-Timestamp
[System.IO.File]::WriteAllText((Join-Path $BundleDir "summary.json"),
    ($summary | ConvertTo-Json -Depth 6), $utf8NoBom)

$report = @(
    $(if ($packageMode) { "Excel regression run of a built package" } else { "Excel regression run" }),
    "",
    "Result:          $($summary.result)",
    "Candidate:       $sha ($branch)",
    "Working tree:    $(if ($dirty) { 'CHANGED (not certification-grade)' } else { 'clean' })",
    "Excel:           $($summary.host.excel_build), $officeBitness",
    "Windows:         $osText",
    "Macro setting:   $($summary.host.macro_policy)",
    "Excel cleanup:   $($summary.cleanup.status)",
    "")
if ($packageMode) {
    $report += @(
        "Package:         $packageName ($packageHost)",
        "Package macros:  $($summary.host.package_macros)",
        "Size:            $($summary.package.bytes) bytes",
        "SHA-256:         $($summary.package.sha256)",
        "The package was opened read-only and never saved. Which commit built it is",
        "recorded by its build evidence, not by this run.",
        "")
    foreach ($key in $packs.Keys) {
        $pack = $summary.packs[$key]
        $report += @(
            "Pack ${key}: $harnessModule.$($pack.entry_point)",
            "  Compile:       $($pack.compile.status) - $($pack.compile.detail)",
            "  Harness:       $($pack.harness.status)$(if ($pack.harness.error) { ' - ' + $pack.harness.error } else { '' })",
            "  Suites:        $(Format-Suites $pack.harness)",
            "  SHA-256 after: $(if ($pack.sha256_after) { $pack.sha256_after } else { '(not read)' })",
            "  Harness summary line, verbatim:",
            "  $(if ($pack.harness.summary_line) { $pack.harness.summary_line } else { '(not found)' })",
            "")
    }
}
else {
    $report += @(
        "Configuration:   $Configuration",
        "Entry point:     $harnessModule.$entryPoint",
        "Import:          $($summary.import.status)",
        "Compile:         $($summary.compile.status) - $($summary.compile.detail)",
        "Harness:         $($summary.harness.status)$(if ($summary.harness.error) { ' - ' + $summary.harness.error } else { '' })",
        "Suites:          $(Format-Suites $summary.harness)",
        "",
        "Harness summary line, verbatim:",
        "$(if ($summary.harness.summary_line) { $summary.harness.summary_line } else { '(not found)' })",
        "")
}
$report += @(
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
