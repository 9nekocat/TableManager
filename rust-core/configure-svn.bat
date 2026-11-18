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

echo Configuring Subversion command-line integration...
"%~dp0table-core.exe" install-svn
set "EXIT_CODE=%errorlevel%"

if not "%EXIT_CODE%"=="0" (
    echo [ERROR] Subversion configuration failed with exit code %EXIT_CODE%.
    if "%QUIET%"=="0" pause
    exit /b %EXIT_CODE%
)

echo Configuring TortoiseSVN integration...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0svn\configure-tortoisesvn.ps1" -RootDirectory "%~dp0."
set "EXIT_CODE=%errorlevel%"

if not "%EXIT_CODE%"=="0" (
    echo [ERROR] TortoiseSVN configuration failed with exit code %EXIT_CODE%.
    if "%QUIET%"=="0" pause
    exit /b %EXIT_CODE%
)

echo.
echo Subversion and TortoiseSVN configuration completed.
if "%QUIET%"=="0" pause
exit /b 0
