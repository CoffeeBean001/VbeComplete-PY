@echo off
rem Launch the C# build. It has NO main window -- it lives in the notification area.
rem
rem v131a: the previous version printed nothing and exited right after `start`,
rem so a double-click looked like "window flashed, nothing started". We now wait
rem and *prove* the process is alive; if it is not, the log tail is printed and
rem the window stays open. A GUI app that dies silently must never fail silently.
rem
rem NOTE (v129): never put a bare parenthesis in echo text -- cmd ends the
rem enclosing if-block at the first ), silently drops the block, exits 255.
setlocal
set "EXE=%~dp0cs\dist\VbeComplete.exe"

if not exist "%EXE%" goto nobuild

start "" "%EXE%"

rem 2s of grace, then check the process list.
ping -n 3 127.0.0.1 >nul 2>nul
rem findstr, not find: Git for Windows ships a Unix `find` that wins on PATH and
rem swallows the /i switch, making this check fail even when the app is alive.
tasklist /fi "imagename eq VbeComplete.exe" /fo csv /nh 2>nul | findstr /i /c:"VbeComplete.exe" >nul
if errorlevel 1 goto died

echo.
echo   VbeComplete C# is running.
echo.
echo   It has no main window. Look for the tray icon in the overflow area
echo   near the clock - you may need to click the small arrow to see it.
echo   Right-click that icon to exit, or run stop-cs.bat.
echo.
ping -n 4 127.0.0.1 >nul 2>nul
endlocal
exit /b 0

:died
echo.
echo [VbeComplete] Launch failed - no process after 2 seconds.
echo ------------------------------------------------------------
if exist "%~dp0crash.log" (
  echo === crash.log ===
  powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0crash.log' -Tail 25" 2>nul
  echo.
)
if exist "%~dp0VbeComplete.log" (
  echo === VbeComplete.log - last 15 lines ===
  powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0VbeComplete.log' -Tail 15" 2>nul
)
echo ------------------------------------------------------------
echo If both sections are empty, set VBECOMPLETE_LOG=1 and run again.
pause
endlocal
exit /b 1

:nobuild
echo.
echo [VbeComplete] cs\dist\VbeComplete.exe not found.
echo Run cs\build.bat first - it compiles and publishes the exe.
echo.
pause
endlocal
exit /b 1
