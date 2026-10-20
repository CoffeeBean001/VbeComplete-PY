@echo off
rem 重新编译 C# 版并发布成单文件 exe（输出到 cs\dist\VbeComplete.exe）
setlocal
cd /d "%~dp0src"
echo [build] 正在编译...
dotnet publish -c Release -r win-x64 --self-contained false -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -o ..\dist
if errorlevel 1 goto fail
echo.
echo [build] 完成：%~dp0dist\VbeComplete.exe
echo [build] 自检中...
"%~dp0dist\VbeComplete.exe" --selftest
echo [build] 自检结果见 selftest.log
goto end
:fail
echo.
echo [build] 失败，看上面的错误信息
:end
endlocal
pause
