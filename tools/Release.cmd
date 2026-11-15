@echo off
setlocal EnableExtensions
set "REPO_ROOT=%~dp0.."
cd /d "%REPO_ROOT%"

set "SCRIPT=%~dp0Release.ps1"

if not exist "%SCRIPT%" (
    echo ERROR: PowerShell script not found:
    echo   %SCRIPT%
    echo.
    pause
    exit /b 1
)

rem Runs the whole release procedure on the checked-out main commit: static
rem gates, source runs, both package builds and runs, the manual PDF, the manual
rem checks, the evidence, then the tag and the GitHub Release. -DryRun stops
rem before the tag. See RELEASING.md. Extra arguments are passed to the script,
rem for example -DryRun -Title "..." -Refs 88,89 -EvidenceIssue 122.
echo Release of the checked-out main commit
echo Close other Excel work first. Answer the questions when asked.
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" %*
set "CODE=%ERRORLEVEL%"
echo.
pause
exit /b %CODE%
