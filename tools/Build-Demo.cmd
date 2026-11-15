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

rem Builds the demo workbook of the checked-out commit with its Ribbon, its demo
rem sheet and a Start sheet, reopens it in Excel for a check and writes an
rem evidence folder next to the checkout (Build-Addin.ps1 -Demo). See
rem tools\README.md. Extra arguments are passed to the script, for example
rem -PackageName DATETIMEPICKER-demo-v1.2.3.xlsm or -AllowDirty.
echo Demo workbook build and check of the checked-out commit
echo Close other Excel work first. Answer the questions when asked.
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -Demo %*
set "CODE=%ERRORLEVEL%"
echo.
pause
exit /b %CODE%
