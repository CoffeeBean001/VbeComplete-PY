@echo off
rem 启动 C# 版。注意：echo 的文本里绝对不要出现裸右括号，cmd 会把它当成
rem if 块的结束符，整块丢弃并返回 255（v129 踩过这个坑）。
setlocal
set "EXE=%~dp0cs\dist\VbeComplete.exe"
if not exist "%EXE%" goto nobuild
start "" "%EXE%"
exit /b 0
:nobuild
echo [VbeComplete] 还没有编译过，请先双击 cs\build.bat
pause
exit /b 1
