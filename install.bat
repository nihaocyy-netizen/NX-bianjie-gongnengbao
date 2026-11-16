@echo off
setlocal
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1" %*
set CODE=%ERRORLEVEL%
echo.
if not "%CODE%"=="0" echo [ERROR] install exit code %CODE%
pause
exit /b %CODE%
