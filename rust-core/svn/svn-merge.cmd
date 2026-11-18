@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0svn-merge.ps1" %*
exit /b %errorlevel%
