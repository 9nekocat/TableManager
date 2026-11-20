@echo off
setlocal
cd /d "%~dp0"

set "QUIET=0"
if /I "%~1"=="/quiet" set "QUIET=1"
if /I "%~1"=="--quiet" set "QUIET=1"

if not exist "table-core.exe" (
    echo [ERROR] table-core.exe was not found next to this script.
    if "%QUIET%"=="0" pause
    exit /b 2
)

if not exist "svn\svn-diff.cmd" (
    echo [ERROR] svn\svn-diff.cmd was not found.
    if "%QUIET%"=="0" pause
    exit /b 2
)

if not exist "svn\svn-merge.cmd" (
    echo [ERROR] svn\svn-merge.cmd was not found.
    if "%QUIET%"=="0" pause
    exit /b 2
)

echo Registering Excel diff and merge tools in TortoiseSVN...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0svn\configure-tortoisesvn.ps1"
set "EXIT_CODE=%errorlevel%"

if not "%EXIT_CODE%"=="0" (
    echo [ERROR] TortoiseSVN configuration failed with exit code %EXIT_CODE%.
    if "%QUIET%"=="0" pause
    exit /b %EXIT_CODE%
)

echo.
echo TortoiseSVN Excel configuration completed.
if "%QUIET%"=="0" pause
exit /b 0
