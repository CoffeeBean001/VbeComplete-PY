@echo off
rem 停止 C# 版
setlocal
taskkill /IM VbeComplete.exe /F >nul 2>&1
if errorlevel 1 goto norun
echo VbeComplete C# stopped.
exit /b 0
:norun
echo No VbeComplete C# process is running.
exit /b 1
