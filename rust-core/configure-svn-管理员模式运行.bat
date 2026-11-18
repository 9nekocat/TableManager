@echo off
setlocal
call "%~dp0configure-svn.bat" %*
exit /b %errorlevel%
