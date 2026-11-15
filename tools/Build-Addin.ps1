<#
.SYNOPSIS
Build the add-in package of the checked-out commit with its Ribbon, reopen it in
Excel to check the Ribbon, and write an evidence folder that names the commit,
the package hash and the host.

.DESCRIPTION
For the commit checked out in this repository:
1. reads the full SHA and refuses a working tree with changes to tracked files,
   unless -AllowDirty is given (the evidence then says it is not
   release-grade);
2. creates the evidence folder outside the checkout and starts session.log;
3. finds Python 3.11 or later and runs .github/scripts/check-ribbon-resources.py
   on src/ribbon, so a missing Ribbon part or image stops the build
   (RELEASING.md section 7, step 5);
4. records Windows, Excel version and build, Office bitness, the macro setting
   and "Trust access to the VBA project object model", as found. It never
   changes any of them;
5. starts a visible Excel instance of its own, adds a new workbook and imports
   the components of one configuration of tools\vba-projects.json
   (packaged-hosts by default, which includes the regression harness, step 7)
   in their listed order. Each component must keep its repository name;
6. writes the configuration's tracked ThisWorkbook code (document_modules in
   tools\vba-projects.json, src\addin\ThisWorkbook.cls) into the workbook's
   own ThisWorkbook module. A document module cannot be imported as a
   component, so its code is copied in, and the file's SHA-256 is recorded;
7. compiles the project with Debug > Compile VBAProject and saves it as an
   .xlam add-in in the evidence folder;
8. injects src/ribbon/ into the package's customUI/ folder with
   tools\build_ribbon_package.py, which also adds the root relationship and
   verifies the result;
9. records the package's size and SHA-256, and runs
   tools\inspect_vba_package.py on it (advisory; its verdict is recorded but
   does not decide the build);
10. opens the package in a fresh Excel instance and asks you three questions
    about the Ribbon: the Show Picker icon, the Reset and Demo icons, and
    whether Show Picker opens the picker. It closes without saving and checks
    that the package hash has not changed. Each Excel instance is closed with
    tools\ExcelCleanup.psm1: every close step is tried even when an earlier one
    fails, and each failure is recorded;
11. writes summary.json and build.txt, and opens the folder.

The exit code is 0 only when every step passed, all three answers were yes,
the package hash was unchanged after reopening, Excel accepted every close
request, and the working tree was clean. Everything else exits 1, and the
evidence says why.

This automates the package build and the Ribbon check of RELEASING.md
section 7. It does not run the regression pack on the package (use
tools\Certify-Excel.cmd for the source, and step 9 of section 7 for packages),
and it builds no demo workbook.

.REQUIREMENTS
- Windows with desktop Excel, Git and Python 3.11 or later on the PATH
  (python, or the py launcher).
- "Trust access to the VBA project object model" already on (Excel > File >
  Options > Trust Center > Trust Center Settings > Macro Settings). The import
  needs it; this script never turns it on.
- Close other Excel work first.

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Build-Addin.ps1

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Build-Addin.ps1 -PackageName DATETIMEPICKER.v1.2.3.xlam
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$BundleDir,
    [Parameter(Mandatory = $false)]
    [string]$Configuration = "packaged-hosts",
    [Parameter(Mandatory = $false)]
    [string]$PackageName,
    [Parameter(Mandatory = $false)]
    [string]$Operator = "Repository owner, local Windows workstation",
    [Parameter(Mandatory = $false)]
    [string]$RepoRoot,
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
$xlOpenXMLAddIn = 55
$vbextDocument = 100
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$script:sessionLog = $null
$script:cleanupFailures = New-Object System.Collections.Generic.List[string]
$windows1252 = [System.Text.Encoding]::GetEncoding(1252)

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

function Invoke-Native {
    # Runs a program and returns its exit code and output lines. Output on
    # stderr must not stop the script, so errors are not escalated here.
    # Python writes UTF-8 and the output is read as UTF-8, so a character
    # outside the console code page neither stops Python nor garbles the log.
    param([string]$Program, [string[]]$Arguments)
    $previous = $ErrorActionPreference
    $previousUtf8 = $env:PYTHONUTF8
    $previousEncoding = $null
    try {
        $previousEncoding = [Console]::OutputEncoding
        [Console]::OutputEncoding = $script:utf8NoBom
    }
    catch {
    }
    $ErrorActionPreference = "Continue"
    $env:PYTHONUTF8 = "1"
    try {
        $output = @(& $Program @Arguments 2>&1 | ForEach-Object { [string]$_ })
        $code = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $previous
        $env:PYTHONUTF8 = $previousUtf8
        if ($null -ne $previousEncoding) {
            try {
                [Console]::OutputEncoding = $previousEncoding
            }
            catch {
            }
        }
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

function Read-YesNo {
    param([string]$Question)
    while ($true) {
        $answer = Read-Host "$Question (Y/N)"
        if ($answer -match '^\s*[yY]') {
            return $true
        }
        if ($answer -match '^\s*[nN]') {
            return $false
        }
    }
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
           "(git status shows none), or pass -AllowDirty to try the tool without release-grade evidence.")
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
$ribbonRoot = Join-Path $RepoRoot "src\ribbon"

# The add-in's ThisWorkbook code is tracked as an exported document module and
# listed under document_modules; it is copied into the workbook, not imported.
$documentModules = @()
$documentNode = $configurationNode.Value.PSObject.Properties | Where-Object { $_.Name -eq "document_modules" }
if ($null -ne $documentNode) {
    $documentModules = @($documentNode.Value)
}
$workbookSources = @($documentModules | Where-Object { [System.IO.Path]::GetFileNameWithoutExtension([string]$_) -eq "ThisWorkbook" })
if ($workbookSources.Count -ne 1) {
    throw "Configuration '$Configuration' must list exactly one ThisWorkbook file under document_modules."
}
$workbookSource = [string]$workbookSources[0]
$workbookSourceFile = Join-Path $RepoRoot ($workbookSource -replace '/', '\')
if (-not (Test-Path -LiteralPath $workbookSourceFile)) {
    throw "The ThisWorkbook source is missing from the checkout: $workbookSource"
}
# Everything after the export header (VERSION, BEGIN...END and Attribute lines)
# is the module's code.
$workbookLines = [System.IO.File]::ReadAllLines($workbookSourceFile, $windows1252)
$first = 0
while ($first -lt $workbookLines.Count -and
    $workbookLines[$first] -match '^(VERSION \d|BEGIN\s*$|END\s*$|\s+MultiUse\s*=|Attribute VB_)') {
    $first++
}
$workbookCode = ""
if ($first -lt $workbookLines.Count) {
    $workbookCode = ($workbookLines[$first..($workbookLines.Count - 1)] -join "`r`n")
}
if ($workbookCode -notmatch '(?m)^\s*Private Sub Workbook_Open\(\)') {
    throw "$workbookSource has no Workbook_Open procedure after its export header."
}
$workbookSourceHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $workbookSourceFile).Hash.ToLowerInvariant()

if ([string]::IsNullOrWhiteSpace($PackageName)) {
    $PackageName = "DATETIMEPICKER-build-" + $sha.Substring(0, 12) + ".xlam"
}
if ($PackageName -notmatch '^[A-Za-z0-9._-]+\.xlam$') {
    throw "-PackageName must be a plain file name ending in .xlam: $PackageName"
}

$stamp = (Get-Date).ToString("yyyyMMdd-HHmmss")
if ([string]::IsNullOrWhiteSpace($BundleDir)) {
    $BundleDir = Join-Path (Split-Path -Parent $RepoRoot) ("DATETIMEPICKER-addin-" + $sha.Substring(0, 12) + "-" + $stamp)
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
$packagePath = Join-Path $BundleDir $PackageName
$startedAt = Get-Timestamp
Write-Session "Add-in build of $sha ($branch)"
Write-Session "Checkout $RepoRoot; evidence $BundleDir; configuration $Configuration; package $PackageName"
if ($dirty) {
    Write-Session "The working tree has changes (-AllowDirty): this build is not release-grade"
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
    tool               = "tools/Build-Addin.ps1"
    operator           = $Operator
    candidate_sha      = $sha
    branch             = $branch
    working_tree_clean = (-not $dirty)
    configuration      = $Configuration
    started_at         = $startedAt
    finished_at        = $null
    host               = [ordered]@{
        os                 = $osText
        excel_version      = "unknown"
        excel_build        = "unknown"
        office_bitness     = $officeBitness
        macro_policy       = "unknown"
        vba_project_access = "unknown"
        python             = "unknown"
    }
    ribbon_source      = [ordered]@{ status = "NOT_RUN"; output = @() }
    import             = [ordered]@{ status = "NOT_RUN"; components = @() }
    lifecycle_hooks    = [ordered]@{ status = "NOT_RUN"; module = $null; source = $workbookSource; source_sha256 = $workbookSourceHash }
    compile            = [ordered]@{ status = "NOT_RUN"; detail = "Not run" }
    save               = [ordered]@{ status = "NOT_RUN" }
    ribbon_injection   = [ordered]@{ status = "NOT_RUN"; output = @() }
    package            = [ordered]@{ file = $PackageName; bytes = $null; sha256 = $null; sha256_after_reopen = $null }
    package_inspection = [ordered]@{ status = "NOT_RUN"; exit_code = $null }
    ribbon_check       = [ordered]@{
        status            = "NOT_RUN"
        show_picker_icon  = $null
        menu_icons        = $null
        show_picker_opens = $null
    }
    cleanup            = [ordered]@{ status = "NOT_RUN"; failures = @() }
    result             = "FAIL"
    reasons            = @()
}

$python = $null
$excel = $null
$workbook = $null
try {
    # ------------------------------------------------------------------------
    # Python and the Ribbon source
    # ------------------------------------------------------------------------
    $python = Find-Python
    if ($null -eq $python) {
        throw "Python 3.11 or later was not found (python or py -3). Install it and run again."
    }
    $summary.host.python = $python.Version
    Write-Session "Python: $($python.Version)"
    $checker = Join-Path $RepoRoot ".github\scripts\check-ribbon-resources.py"
    $check = Invoke-Python -Python $python -Arguments @($checker, "--root", $ribbonRoot)
    $summary.ribbon_source = [ordered]@{ status = $(if ($check.Code -eq 0) { "PASS" } else { "FAIL" }); output = $check.Output }
    foreach ($line in $check.Output) {
        Write-Session "  $line"
    }
    if ($check.Code -ne 0) {
        throw "The Ribbon source is incomplete; see check-ribbon-resources.py above."
    }

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
    # Lifecycle hooks in ThisWorkbook
    # ------------------------------------------------------------------------
    # The workbook's document module has a localized name in some Excel
    # languages, and a workbook created by code can report an empty CodeName.
    # Only the workbook's own document module exposes an IsAddin property, so
    # that identifies it; the CodeName is the fallback.
    $documents = @()
    foreach ($candidate in $project.VBComponents) {
        if ([int]$candidate.Type -eq $vbextDocument) {
            $documents += $candidate
        }
    }
    $thisWorkbook = @()
    foreach ($document in $documents) {
        try {
            $null = $document.Properties.Item("IsAddin")
            $thisWorkbook += $document
        }
        catch {
        }
    }
    if ($thisWorkbook.Count -ne 1) {
        $workbookCodeName = [string]$workbook.CodeName
        $thisWorkbook = @($documents | Where-Object {
                -not [string]::IsNullOrEmpty($workbookCodeName) -and [string]$_.Name -eq $workbookCodeName })
    }
    if ($thisWorkbook.Count -ne 1) {
        $summary.lifecycle_hooks.status = "FAIL"
        throw "The workbook's ThisWorkbook module could not be identified."
    }
    $codeModule = $thisWorkbook[0].CodeModule
    if ($codeModule.CountOfLines -gt 0) {
        $codeModule.DeleteLines(1, $codeModule.CountOfLines)
    }
    $codeModule.AddFromString($workbookCode)
    $summary.lifecycle_hooks.status = "PASS"
    $summary.lifecycle_hooks.module = [string]$thisWorkbook[0].Name
    Write-Session "Wrote $workbookSource (SHA-256 $workbookSourceHash) into $($thisWorkbook[0].Name)"

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
        Write-Host ""
        Write-Host "Debug > Compile VBAProject still looks enabled. In the VBE window, choose"
        Write-Host "Debug > Compile VBAProject yourself. Answer Y if it completes with no error and"
        Write-Host "the command is then greyed out; answer N if it reports a compile error."
        if (-not (Read-YesNo "Compiled with no error?")) {
            $summary.compile = [ordered]@{ status = "FAIL"; detail = "Debug > Compile VBAProject reported an error, as the operator saw" }
            throw "The project does not compile; see the VBE."
        }
        $compileDetail = "Debug > Compile VBAProject completed with no error, as the operator saw in the VBE"
    }
    $summary.compile = [ordered]@{ status = "PASS"; detail = $compileDetail }
    Write-Session $compileDetail
    $vbe.MainWindow.Visible = $false

    # ------------------------------------------------------------------------
    # Save as an add-in and close the build instance
    # ------------------------------------------------------------------------
    $workbook.SaveAs($packagePath, $xlOpenXMLAddIn)
    $summary.save = [ordered]@{ status = "PASS" }
    Write-Session "Saved as $packagePath"
    # Workbook_BeforeClose would call DP_Stop on a runtime this instance never
    # started; events are off for the close.
    $excel.EnableEvents = $false
    Close-Excel -Excel $excel -Workbooks @($workbook) -Label "Build"
    $workbook = $null
    $excel = $null

    # ------------------------------------------------------------------------
    # Inject the Ribbon and identify the package
    # ------------------------------------------------------------------------
    $injector = Join-Path $RepoRoot "tools\build_ribbon_package.py"
    $injection = Invoke-Python -Python $python -Arguments @(
        $injector, "--package", $packagePath, "--ribbon-root", $ribbonRoot,
        "--report", (Join-Path $BundleDir "ribbon-injection.json"))
    $summary.ribbon_injection = [ordered]@{ status = $(if ($injection.Code -eq 0) { "PASS" } else { "FAIL" }); output = $injection.Output }
    foreach ($line in $injection.Output) {
        Write-Session "  $line"
    }
    if ($injection.Code -ne 0) {
        throw "The Ribbon could not be injected; see build_ribbon_package.py above."
    }
    $summary.package.bytes = (Get-Item -LiteralPath $packagePath).Length
    $summary.package.sha256 = (Get-FileHash -Algorithm SHA256 -LiteralPath $packagePath).Hash.ToLowerInvariant()
    Write-Session "Package $PackageName, $($summary.package.bytes) bytes, SHA-256 $($summary.package.sha256)"

    $inspector = Join-Path $RepoRoot "tools\inspect_vba_package.py"
    $inspection = Invoke-Python -Python $python -Arguments @(
        $inspector, "--package", $packagePath, "--source", $RepoRoot, "--source-rev", $sha,
        "--configuration", $Configuration,
        "--output", (Join-Path $BundleDir "package-inspection.json"),
        "--summary", (Join-Path $BundleDir "package-inspection.md"))
    $summary.package_inspection = [ordered]@{
        status    = $(if ($inspection.Code -eq 0) { "PASS" } else { "FINDINGS" })
        exit_code = $inspection.Code
    }
    Write-Session "Package inspection (advisory): $($summary.package_inspection.status), exit code $($inspection.Code); see package-inspection.md"

    # ------------------------------------------------------------------------
    # Reopen the package and check the Ribbon
    # ------------------------------------------------------------------------
    $summary.ribbon_check.status = "STARTED"
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $addin = $excel.Workbooks.Open($packagePath, 0, $true)
    $blank = $excel.Workbooks.Add()
    try {
        $excel.WindowState = -4137
    }
    catch {
    }
    Write-Session "Reopened $PackageName read-only in a new Excel instance, with a blank workbook"
    Write-Host ""
    Write-Host "Switch to the new Excel window and open its Home tab."
    $summary.ribbon_check.show_picker_icon = Read-YesNo "1. The 'Date / Time Picker' group shows a large 'Show Picker' button with its calendar-grid icon?"
    $summary.ribbon_check.menu_icons = Read-YesNo "2. The arrow under 'Show Picker' opens a menu where 'Reset' and 'Demo' each show their own icon?"
    $summary.ribbon_check.show_picker_opens = Read-YesNo "3. Clicking 'Show Picker' opens the Date / Time Picker? (close it again before answering)"
    $summary.ribbon_check.status = "ANSWERED"
    Write-Session ("Ribbon check: Show Picker icon=" + $summary.ribbon_check.show_picker_icon +
                   "; Reset and Demo icons=" + $summary.ribbon_check.menu_icons +
                   "; Show Picker opens the picker=" + $summary.ribbon_check.show_picker_opens)
    Close-Excel -Excel $excel -Workbooks @($blank, $addin) -Label "Check"
    $excel = $null
    $summary.package.sha256_after_reopen = (Get-FileHash -Algorithm SHA256 -LiteralPath $packagePath).Hash.ToLowerInvariant()
    Write-Session "SHA-256 after reopening: $($summary.package.sha256_after_reopen)"
}
catch {
    $summary.reasons += ("Stopped: " + $_.Exception.Message)
    Write-Session ("Stopped: " + $_.Exception.Message)
}
finally {
    if ($null -ne $excel) {
        try {
            $excel.EnableEvents = $false
        }
        catch {
        }
        $open = @()
        try {
            foreach ($book in $excel.Workbooks) {
                $open += $book
            }
        }
        catch {
        }
        Close-Excel -Excel $excel -Workbooks $open -Label "Remaining"
    }
}

# ----------------------------------------------------------------------------
# Verdict and evidence
# ----------------------------------------------------------------------------
Add-CleanupResult -Summary $summary -Failures $script:cleanupFailures.ToArray()
if ($dirty) { $summary.reasons += "The working tree had changes to tracked files" }
if ($summary.ribbon_source.status -ne "PASS") { $summary.reasons += "Ribbon source: $($summary.ribbon_source.status)" }
if ($summary.import.status -ne "PASS") { $summary.reasons += "Import: $($summary.import.status)" }
if ($summary.lifecycle_hooks.status -ne "PASS") { $summary.reasons += "ThisWorkbook code: $($summary.lifecycle_hooks.status)" }
if ($summary.compile.status -ne "PASS") { $summary.reasons += "Compile: $($summary.compile.status)" }
if ($summary.save.status -ne "PASS") { $summary.reasons += "Save as add-in: $($summary.save.status)" }
if ($summary.ribbon_injection.status -ne "PASS") { $summary.reasons += "Ribbon injection: $($summary.ribbon_injection.status)" }
if ($summary.ribbon_check.status -ne "ANSWERED") {
    $summary.reasons += "Ribbon check: $($summary.ribbon_check.status)"
}
else {
    if (-not $summary.ribbon_check.show_picker_icon) { $summary.reasons += "The Show Picker icon did not render" }
    if (-not $summary.ribbon_check.menu_icons) { $summary.reasons += "The Reset or Demo icon did not render" }
    if (-not $summary.ribbon_check.show_picker_opens) { $summary.reasons += "Show Picker did not open the picker" }
}
if ($null -ne $summary.package.sha256 -and $summary.package.sha256_after_reopen -ne $summary.package.sha256) {
    $summary.reasons += "The package hash changed after reopening, or could not be read"
}
if ($summary.reasons.Count -eq 0) {
    $summary.result = "PASS"
}
$summary.finished_at = Get-Timestamp
[System.IO.File]::WriteAllText((Join-Path $BundleDir "summary.json"),
    ($summary | ConvertTo-Json -Depth 6), $utf8NoBom)

function Format-Answer {
    param([object]$Value)
    if ($null -eq $Value) { return "not asked" }
    if ($Value) { return "yes" }
    return "no"
}

$report = @(
    "Add-in build and Ribbon check",
    "",
    "Result:          $($summary.result)",
    "Candidate:       $sha ($branch)",
    "Working tree:    $(if ($dirty) { 'CHANGED (not release-grade)' } else { 'clean' })",
    "Configuration:   $Configuration",
    "Excel:           $($summary.host.excel_build), $officeBitness",
    "Windows:         $osText",
    "Macro setting:   $($summary.host.macro_policy)",
    "Python:          $($summary.host.python)",
    "Ribbon source:   $($summary.ribbon_source.status)",
    "Import:          $($summary.import.status)",
    "ThisWorkbook:    $($summary.lifecycle_hooks.status) - $workbookSource into $($summary.lifecycle_hooks.module)",
    "                 source SHA-256 $workbookSourceHash",
    "Compile:         $($summary.compile.status) - $($summary.compile.detail)",
    "Save as .xlam:   $($summary.save.status)",
    "Ribbon inject:   $($summary.ribbon_injection.status)",
    "Excel cleanup:   $($summary.cleanup.status)",
    "",
    "Package:         $PackageName",
    "Size:            $($summary.package.bytes) bytes",
    "SHA-256:         $($summary.package.sha256)",
    "After reopening: $($summary.package.sha256_after_reopen)",
    "Inspection:      $($summary.package_inspection.status) (advisory, exit code $($summary.package_inspection.exit_code))",
    "",
    "Ribbon check, as the operator saw it:",
    "  Show Picker shows its icon:        $(Format-Answer $summary.ribbon_check.show_picker_icon)",
    "  Reset and Demo show their icons:   $(Format-Answer $summary.ribbon_check.menu_icons)",
    "  Show Picker opens the picker:      $(Format-Answer $summary.ribbon_check.show_picker_opens)",
    "",
    "Started:  $startedAt",
    "Finished: $($summary.finished_at)")
if ($summary.reasons.Count -gt 0) {
    $report += @("", "Not passing because:")
    foreach ($reason in $summary.reasons) {
        $report += "- $reason"
    }
}
[System.IO.File]::WriteAllLines((Join-Path $BundleDir "build.txt"), $report, $utf8NoBom)

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
