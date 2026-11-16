@echo off
setlocal
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1" %*
set CODE=%ERRORLEVEL%
echo.
if not "%CODE%"=="0" echo [ERROR] uninstall exit code %CODE%
pause
exit /b %CODE%
