@echo off
setlocal
cd /d "%~dp0"

if not exist "table-core.exe" (
    echo [ERROR] table-core.exe was not found next to this script.
    pause
    exit /b 2
)

if not exist "svn\svn-diff.cmd" (
    echo [ERROR] svn\svn-diff.cmd was not found.
    pause
    exit /b 2
)

if not exist "svn\svn-merge.cmd" (
    echo [ERROR] svn\svn-merge.cmd was not found.
    pause
    exit /b 2
)

echo Configuring Subversion to use TableManager for xlsx-family files...
"%~dp0table-core.exe" install-svn
set "EXIT_CODE=%errorlevel%"

if not "%EXIT_CODE%"=="0" (
    echo [ERROR] SVN configuration failed with exit code %EXIT_CODE%.
    pause
    exit /b %EXIT_CODE%
)

echo.
echo SVN configuration completed.
pause
exit /b 0
