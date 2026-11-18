@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0svn-diff.ps1" %*
exit /b %errorlevel%
