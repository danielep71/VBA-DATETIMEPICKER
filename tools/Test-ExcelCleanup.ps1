<#
.SYNOPSIS
Check, without Excel, that a failed Excel cleanup in Certify-Excel.ps1 or
Build-Addin.ps1 is recorded and cannot give a passing verdict.

.DESCRIPTION
Runs tools/ExcelCleanup.psm1 against stand-in workbook and Excel objects whose
Close and Quit succeed or throw, and checks that:
- every step is tried even when an earlier one throws;
- each failure is returned, and none is thrown;
- Add-CleanupResult records the failures as reasons, so a summary that would
  otherwise pass ends with FAIL, as the scripts' verdict decides.

Prints one line per case and exits 0 only when every case passes. Runs on
Windows PowerShell 5.1 and PowerShell 7.

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Test-ExcelCleanup.ps1
#>
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Import-Module (Join-Path $PSScriptRoot "ExcelCleanup.psm1") -Force

$script:failed = 0

function New-StandIn {
    # An object with one method that records its call and throws on request.
    # With an HResult it throws a COMException carrying it, as a dead Excel does.
    param([string]$Method, [bool]$Throws, [System.Collections.Generic.List[string]]$Calls, [string]$Name,
          [int]$HResult = 0)
    $object = [pscustomobject]@{ Name = $Name; Throws = $Throws; Calls = $Calls; HResult = $HResult }
    $body = {
        $this.Calls.Add($this.Name)
        if ($this.HResult -ne 0) {
            throw (New-Object System.Runtime.InteropServices.COMException("$($this.Name) gone", $this.HResult))
        }
        if ($this.Throws) {
            throw "$($this.Name) refused"
        }
    }
    $object | Add-Member -MemberType ScriptMethod -Name $Method -Value $body
    return $object
}

function Test-Case {
    param([string]$Name, [scriptblock]$Body)
    try {
        & $Body
        Write-Host "PASS  $Name"
    }
    catch {
        $script:failed++
        Write-Host "FAIL  $Name - $($_.Exception.Message)"
    }
}

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) {
        throw $Message
    }
}

function Invoke-Cleanup {
    param([bool]$BookThrows, [bool]$QuitThrows, [int]$HResult = 0)
    $calls = New-Object System.Collections.Generic.List[string]
    $book = New-StandIn -Method "Close" -Throws $BookThrows -Calls $calls -Name "workbook" -HResult $HResult
    $excel = New-StandIn -Method "Quit" -Throws $QuitThrows -Calls $calls -Name "excel" -HResult $HResult
    $logged = New-Object System.Collections.Generic.List[string]
    $log = { param([string]$Text) $logged.Add($Text) }.GetNewClosure()
    $failures = @(Close-ExcelInstance -Excel $excel -Workbooks @($book) -Label "Run" -Log $log)
    return [pscustomobject]@{ Calls = $calls; Failures = $failures; Logged = $logged }
}

function New-PassingSummary {
    # A summary in which every other check passed.
    return [ordered]@{ result = "FAIL"; reasons = @() }
}

function Get-Verdict {
    # The verdict rule of Certify-Excel.ps1 and Build-Addin.ps1.
    param([System.Collections.IDictionary]$Summary)
    if (@($Summary.reasons).Count -eq 0) {
        return "PASS"
    }
    return "FAIL"
}

Test-Case "clean cleanup returns no failure and passes" {
    $run = Invoke-Cleanup -BookThrows $false -QuitThrows $false
    Assert-True ($run.Failures.Count -eq 0) "expected no failure, got: $($run.Failures -join '; ')"
    Assert-True (($run.Calls -join ",") -eq "workbook,excel") "expected workbook then excel, got $($run.Calls -join ',')"
    $summary = New-PassingSummary
    Add-CleanupResult -Summary $summary -Failures $run.Failures
    Assert-True ($summary.cleanup.status -eq "PASS") "cleanup status is $($summary.cleanup.status)"
    Assert-True ((Get-Verdict $summary) -eq "PASS") "a clean cleanup must not fail the run"
}

Test-Case "a workbook that will not close is recorded and Excel is still asked to quit" {
    $run = Invoke-Cleanup -BookThrows $true -QuitThrows $false
    Assert-True ($run.Failures.Count -eq 1) "expected 1 failure, got $($run.Failures.Count)"
    Assert-True ($run.Failures[0] -like "Run: a workbook could not be closed: *workbook refused*") "unexpected text: $($run.Failures[0])"
    Assert-True (($run.Calls -join ",") -eq "workbook,excel") "Quit was not tried after the failed Close"
    Assert-True ($run.Logged -contains $run.Failures[0]) "the failure was not written to the log"
}

Test-Case "an Excel that will not quit gives FAIL" {
    $run = Invoke-Cleanup -BookThrows $false -QuitThrows $true
    Assert-True ($run.Failures.Count -eq 1) "expected 1 failure, got $($run.Failures.Count)"
    Assert-True ($run.Failures[0] -like "Run: Excel could not be asked to quit: *excel refused*") "unexpected text: $($run.Failures[0])"
    $summary = New-PassingSummary
    Add-CleanupResult -Summary $summary -Failures $run.Failures
    Assert-True ($summary.cleanup.status -eq "FAIL") "cleanup status is $($summary.cleanup.status)"
    Assert-True (@($summary.reasons).Count -eq 1) "expected 1 reason, got $(@($summary.reasons).Count)"
    Assert-True ((Get-Verdict $summary) -eq "FAIL") "a failed quit must not pass"
}

Test-Case "both failures are recorded and an earlier reason is kept" {
    $run = Invoke-Cleanup -BookThrows $true -QuitThrows $true
    Assert-True ($run.Failures.Count -eq 2) "expected 2 failures, got $($run.Failures.Count)"
    $summary = [ordered]@{ result = "FAIL"; reasons = @("Stopped: the harness raised an error") }
    Add-CleanupResult -Summary $summary -Failures $run.Failures
    Assert-True (@($summary.reasons).Count -eq 3) "expected 3 reasons, got $(@($summary.reasons).Count)"
    Assert-True ($summary.reasons[0] -eq "Stopped: the harness raised an error") "the original error was not kept first"
    Assert-True ($summary.reasons[2] -like "Excel cleanup: Run: Excel could not be asked to quit: *") "unexpected reason: $($summary.reasons[2])"
}

Test-Case "an Excel that has already exited is logged, not counted" {
    # 0x800706BE (the call failed) and 0x800706BA (server unavailable).
    foreach ($code in @(-2147023170, -2147023174)) {
        $run = Invoke-Cleanup -BookThrows $false -QuitThrows $false -HResult $code
        Assert-True ($run.Failures.Count -eq 0) "expected no failure for $code, got: $($run.Failures -join '; ')"
        Assert-True ($run.Logged -contains "Run: Excel had already exited") "the exit was not logged for $code"
        Assert-True (($run.Calls -join ",") -eq "workbook,excel") "Quit was not tried for $code"
    }
}

Test-Case "nothing to close returns no failure" {
    $failures = @(Close-ExcelInstance -Excel $null -Workbooks @($null) -Label "Run")
    Assert-True ($failures.Count -eq 0) "expected no failure, got: $($failures -join '; ')"
    $summary = New-PassingSummary
    Add-CleanupResult -Summary $summary -Failures $failures
    Assert-True ((Get-Verdict $summary) -eq "PASS") "nothing to close must not fail the run"
}

Write-Host ""
if ($script:failed -gt 0) {
    Write-Host "$script:failed case(s) failed"
    exit 1
}
Write-Host "All cases passed"
exit 0
