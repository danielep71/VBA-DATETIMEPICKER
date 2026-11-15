@echo off
setlocal EnableExtensions
set "REPO_ROOT=%~dp0.."
cd /d "%REPO_ROOT%"

set "SCRIPT=%~dp0Build-Addin.ps1"

if not exist "%SCRIPT%" (
    echo ERROR: PowerShell script not found:
    echo   %SCRIPT%
    echo.
    pause
    exit /b 1
)

rem Builds the add-in package of the checked-out commit with its Ribbon, reopens
rem it in Excel for a Ribbon check and writes an evidence folder next to the
rem checkout. See tools\README.md. Extra arguments are passed to the script, for
rem example -PackageName DATETIMEPICKER.v1.2.3.xlam or -AllowDirty.
echo Add-in build and Ribbon check of the checked-out commit
echo Close other Excel work first. Answer the Ribbon questions when asked.
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" %*
set "CODE=%ERRORLEVEL%"
echo.
pause
exit /b %CODE%
