@echo off
rem VbeComplete launcher (silent background, no console window)
rem
rem v128: the entry point is now watchdog.py, not main.py -- the supervisor is
rem what starts the main program. That way a plain double-click also gets
rem crash/hang recovery: when the main thread is stuck in a COM call (process
rem alive, single-instance mutex still held) the supervisor kills it and starts
rem a fresh one, so a reboot is no longer needed.
setlocal
set "HERE=%~dp0"

where py >nul 2>nul
if %errorlevel%==0 (
  set "PY=py -3"
  set "PYW=pyw -3"
) else (
  set "PY=python"
  set "PYW=pythonw"
)

rem Skip pip when the dependencies are already importable -- running it on every
rem launch used to hit the network and made startup look like "nothing happens".
%PY% -c "import pynput, win32gui" >nul 2>nul
if errorlevel 1 (
  echo Installing dependencies - first run only...
  %PY% -m pip install -q -r "%HERE%requirements.txt"
  if errorlevel 1 (
    echo pip install failed. Please run "python main.py" in a console to see errors.
    pause
    exit /b 1
  )
)

echo Starting VbeComplete - runs in background...
start "" %PYW% "%HERE%watchdog.py"
echo Done. You can close this window. The tool stays running.
rem v129: ping instead of `timeout` -- `timeout` refuses to run when stdin is
rem redirected and some setups have a *different* `timeout` earlier in PATH
rem (git's coreutils), which printed "invalid time interval".
ping -n 3 127.0.0.1 >nul 2>nul
endlocal
