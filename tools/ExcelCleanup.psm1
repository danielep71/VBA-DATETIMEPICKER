<#
.SYNOPSIS
Close an Excel instance that a tools script started, and record every step that
failed so that the script's verdict can count it.

.DESCRIPTION
Certify-Excel.ps1 and Build-Addin.ps1 start their own Excel instances and must
close them again whatever happened before. A failure while closing is evidence:
Excel may still be running, a workbook may still be open, or the package file
may still be locked. Such a failure must not produce a passing verdict.

Close-ExcelInstance tries every step even when an earlier one fails: it closes
each workbook without saving, asks Excel to quit and releases each COM
reference. It never throws, so an error that stopped the run before it is kept.
It returns one text per failed step; no text means that Excel accepted every
request. That does not prove that the EXCEL.EXE process has ended: Windows ends
it once every reference to it is released, and a script cannot observe that
reliably.

Add-CleanupResult records the returned texts in a run summary: a cleanup entry
with status PASS or FAIL, and one reason per failure. Both scripts give a PASS
verdict only when the summary has no reason, so a failed cleanup always gives
FAIL and exit code 1.

One case is not a failure: a call refused because the Excel process has
already exited (RPC server unavailable, the remote procedure call failed, or
the object disconnected from its client). Nothing can still be open or locked
then, so it is logged and not counted. It happens when Excel ends by itself
while the last workbook closes, after every test has run.

tools/Test-ExcelCleanup.ps1 checks both functions without Excel.
#>
Set-StrictMode -Version Latest

# RPC_S_SERVER_UNAVAILABLE, RPC_S_CALL_FAILED and RPC_E_DISCONNECTED: the
# Excel process the call was for no longer exists.
$script:ExcelExitedCodes = @(-2147023174, -2147023170, -2147417848)

function Test-ExcelExited {
    param([System.Exception]$Exception)
    for ($current = $Exception; $null -ne $current; $current = $current.InnerException) {
        if ($script:ExcelExitedCodes -contains $current.HResult) {
            return $true
        }
    }
    return $false
}

function Close-ExcelInstance {
    [CmdletBinding()]
    param(
        [object]$Excel,
        [object[]]$Workbooks = @(),
        [string]$Label = "Excel",
        [scriptblock]$Log
    )
    $failures = New-Object System.Collections.Generic.List[string]
    foreach ($book in @($Workbooks)) {
        if ($null -eq $book) {
            continue
        }
        try {
            [void]$book.Close($false)
        }
        catch {
            if (Test-ExcelExited $_.Exception) {
                if ($null -ne $Log) {
                    & $Log "${Label}: Excel had already exited; nothing is left open"
                }
            }
            else {
                $failures.Add("${Label}: a workbook could not be closed: " + $_.Exception.Message)
            }
        }
        try {
            if ([System.Runtime.InteropServices.Marshal]::IsComObject($book)) {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($book)
            }
        }
        catch {
            $failures.Add("${Label}: a workbook reference could not be released: " + $_.Exception.Message)
        }
    }
    if ($null -ne $Excel) {
        try {
            [void]$Excel.Quit()
            if ($null -ne $Log) {
                & $Log "${Label}: Excel was asked to quit"
            }
        }
        catch {
            if (Test-ExcelExited $_.Exception) {
                if ($null -ne $Log) {
                    & $Log "${Label}: Excel had already exited"
                }
            }
            else {
                $failures.Add("${Label}: Excel could not be asked to quit: " + $_.Exception.Message)
            }
        }
        try {
            if ([System.Runtime.InteropServices.Marshal]::IsComObject($Excel)) {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel)
            }
        }
        catch {
            $failures.Add("${Label}: the Excel reference could not be released: " + $_.Exception.Message)
        }
    }
    try {
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
    catch {
        $failures.Add("${Label}: the COM references could not be collected: " + $_.Exception.Message)
    }
    if ($null -ne $Log) {
        foreach ($failure in $failures) {
            & $Log $failure
        }
    }
    return $failures.ToArray()
}

function Add-CleanupResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$Summary,
        [string[]]$Failures = @()
    )
    $list = @($Failures | Where-Object { -not [string]::IsNullOrEmpty($_) })
    $Summary["cleanup"] = [ordered]@{
        status   = $(if ($list.Count -eq 0) { "PASS" } else { "FAIL" })
        failures = $list
    }
    foreach ($failure in $list) {
        $Summary["reasons"] = @($Summary["reasons"]) + ("Excel cleanup: " + $failure)
    }
}

Export-ModuleMember -Function Close-ExcelInstance, Add-CleanupResult
