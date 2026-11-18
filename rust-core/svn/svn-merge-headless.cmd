@echo off
setlocal
set "TM_SVN_HEADLESS=1"
set "TM_SVN_NO_OPEN=1"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0svn-merge.ps1" %*
exit /b %errorlevel%
