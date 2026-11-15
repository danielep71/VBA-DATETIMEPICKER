@echo off
setlocal EnableExtensions
set "REPO_ROOT=%~dp0.."
cd /d "%REPO_ROOT%"

set "SCRIPT=%~dp0Certify-Excel.ps1"

if not exist "%SCRIPT%" (
    echo ERROR: PowerShell script not found:
    echo   %SCRIPT%
    echo.
    pause
    exit /b 1
)

rem Builds a new workbook from the checked-out commit, compiles it, runs the
rem regression pack and writes an evidence folder next to the checkout.
rem -Package <file> runs both packs inside a built .xlam or .xlsm instead.
rem See tools\README.md. Extra arguments are passed to the script, for example
rem -WithUISmoke, -Package <file> or -AllowDirty.
echo Excel regression run of the checked-out commit
echo Close other Excel work first. Do not use Excel until the run ends.
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" %*
set "CODE=%ERRORLEVEL%"
echo.
pause
exit /b %CODE%
