@echo off
rem Rebuild the C# build and publish it as a single-file exe (cs\dist\VbeComplete.exe).
rem Remember: keep echo text free of bare parentheses, and write this file with CRLF.
setlocal
cd /d "%~dp0src"
echo [build] compiling...
dotnet publish -c Release -r win-x64 --self-contained false -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -o ..\dist
if errorlevel 1 goto fail
echo.
echo [build] done: %~dp0dist\VbeComplete.exe
echo [build] running self test...
"%~dp0dist\VbeComplete.exe" --selftest
echo [build] self test log: project root \selftest.log
goto end
:fail
echo.
echo [build] FAILED - see the errors above
:end
endlocal
pause
